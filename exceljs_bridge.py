import atexit
import json
import logging
import math
import os
import queue
import subprocess
import threading
import time
from pathlib import Path
from typing import Any, Dict, Optional


def _sanitize_for_json(obj):
    """Replace NaN/Infinity with None so JSON.parse on the Node side doesn't choke."""
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_for_json(v) for v in obj]
    return obj


class ExcelJSBridgeError(RuntimeError):
    pass


class ExcelJSBridge:
    def __init__(
        self,
        logger: Optional[logging.Logger] = None,
        script_path: Optional[str] = None,
        timeout_seconds: int = 120,
    ) -> None:
        self._logger = logger or logging.getLogger(__name__)
        self._timeout_seconds = timeout_seconds
        self._script_path = script_path or str(Path(__file__).resolve().parent / "excel_bridge" / "excel_operations.js")
        self._process: Optional[subprocess.Popen] = None
        self._request_id = 0
        self._lock = threading.RLock()
        self._responses: Dict[int, queue.Queue] = {}
        self._stderr_buffer: list[str] = []
        self._stop_event = threading.Event()
        self._stdout_thread: Optional[threading.Thread] = None
        self._stderr_thread: Optional[threading.Thread] = None

    def _log(self, level: int, message: str) -> None:
        try:
            self._logger.log(level, message)
        except Exception:
            pass

    def _node_command(self) -> list[str]:
        node_cmd = os.environ.get("EXCELJS_NODE_BIN", "node")
        return [node_cmd, self._script_path]

    def _ensure_process(self) -> None:
        needs_ping = False
        with self._lock:
            if self._process and self._process.poll() is None:
                return

            self._cleanup_process_locked()

            script = Path(self._script_path)
            if not script.exists():
                raise ExcelJSBridgeError(f"ExcelJS bridge script not found: {script}")

            cmd = self._node_command()
            self._log(logging.INFO, f"Starting ExcelJS bridge process: {' '.join(cmd)}")

            self._stop_event.clear()
            self._process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
            )

            self._stdout_thread = threading.Thread(target=self._stdout_reader, daemon=True)
            self._stderr_thread = threading.Thread(target=self._stderr_reader, daemon=True)
            self._stdout_thread.start()
            self._stderr_thread.start()
            needs_ping = True

        if needs_ping:
            self.ping()

    def _stdout_reader(self) -> None:
        process = self._process
        if process is None or process.stdout is None:
            return

        for line in process.stdout:
            if self._stop_event.is_set():
                break
            line = line.strip()
            if not line:
                continue
            try:
                payload = json.loads(line)
            except json.JSONDecodeError as exc:
                self._log(logging.ERROR, f"ExcelJS bridge returned non-JSON line: {line} ({exc})")
                continue

            response_id = payload.get("id")
            if response_id is None:
                self._log(logging.WARNING, f"ExcelJS bridge response missing id: {payload}")
                continue

            with self._lock:
                response_queue = self._responses.get(response_id)
            if response_queue is not None:
                response_queue.put(payload)
            else:
                self._log(logging.WARNING, f"Received response for unknown id {response_id}: {payload}")

    def _stderr_reader(self) -> None:
        process = self._process
        if process is None or process.stderr is None:
            return

        for line in process.stderr:
            if self._stop_event.is_set():
                break
            cleaned = line.rstrip("\n")
            with self._lock:
                self._stderr_buffer.append(cleaned)
                if len(self._stderr_buffer) > 200:
                    self._stderr_buffer = self._stderr_buffer[-200:]
            self._log(logging.INFO, f"[exceljs-bridge] {cleaned}")

    def _next_id(self) -> int:
        with self._lock:
            self._request_id += 1
            return self._request_id

    def _send_request(self, method: str, params: Dict[str, Any], timeout_seconds: Optional[int] = None) -> Dict[str, Any]:
        self._ensure_process()

        with self._lock:
            process = self._process
            if process is None or process.poll() is not None or process.stdin is None:
                raise ExcelJSBridgeError("ExcelJS bridge process is not running")

            request_id = self._next_id()
            response_queue: queue.Queue = queue.Queue(maxsize=1)
            self._responses[request_id] = response_queue

            payload = {
                "jsonrpc": "2.0",
                "method": method,
                "params": params,
                "id": request_id,
            }

            try:
                process.stdin.write(json.dumps(_sanitize_for_json(payload)) + "\n")
                process.stdin.flush()
            except Exception as exc:
                self._responses.pop(request_id, None)
                self._restart_process_locked(reason=f"stdin write failed for method {method}: {exc}")
                raise ExcelJSBridgeError(f"Failed to send request to ExcelJS bridge: {exc}") from exc

        timeout = timeout_seconds or self._timeout_seconds
        try:
            response = response_queue.get(timeout=timeout)
        except queue.Empty as exc:
            with self._lock:
                self._responses.pop(request_id, None)
                recent_stderr = "\n".join(self._stderr_buffer[-30:])
                self._restart_process_locked(reason=f"timeout waiting for {method}")
            raise ExcelJSBridgeError(
                f"Timeout waiting for ExcelJS bridge response for method '{method}' after {timeout}s. "
                f"Recent stderr:\n{recent_stderr}"
            ) from exc

        with self._lock:
            self._responses.pop(request_id, None)

        if "error" in response:
            error = response["error"]
            raise ExcelJSBridgeError(
                f"ExcelJS bridge error in method '{method}': {error.get('message')}\n{error.get('data', '')}"
            )

        return response.get("result", {})

    def _restart_process_locked(self, reason: str) -> None:
        self._log(logging.WARNING, f"Restarting ExcelJS bridge process: {reason}")
        self._cleanup_process_locked()

    def _cleanup_process_locked(self) -> None:
        self._stop_event.set()
        process = self._process

        if process is not None:
            try:
                if process.stdin:
                    process.stdin.close()
            except Exception:
                pass

            if process.poll() is None:
                try:
                    process.terminate()
                    process.wait(timeout=2)
                except Exception:
                    try:
                        process.kill()
                    except Exception:
                        pass

        self._process = None

        for response_queue in self._responses.values():
            try:
                response_queue.put_nowait({
                    "jsonrpc": "2.0",
                    "error": {"message": "Process terminated"},
                    "id": None,
                })
            except Exception:
                pass
        self._responses.clear()

    def close(self) -> None:
        with self._lock:
            self._cleanup_process_locked()

    def ping(self) -> Dict[str, Any]:
        return self._send_request("ping", {}, timeout_seconds=10)

    def write_excel_distro(
        self,
        template_path: str,
        temp_dir: str,
        image_data: list[dict],
        header_row: int,
        row_offset: int = 0,
    ) -> Dict[str, Any]:
        return self._send_request(
            "writeExcelDistro",
            {
                "templatePath": template_path,
                "tempDir": temp_dir,
                "imageData": image_data,
                "headerRow": int(header_row),
                "rowOffset": int(row_offset),
            },
        )

    def write_excel_msrp(
        self,
        template_path: str,
        temp_dir: str,
        image_data: list[dict],
        header_row: int,
        target_column: str,
        row_offset: int,
        populate_images: bool = True,
        populate_msrp: bool = True,
    ) -> Dict[str, Any]:
        return self._send_request(
            "writeExcelMSRP",
            {
                "templatePath": template_path,
                "tempDir": temp_dir,
                "imageData": image_data,
                "headerRow": int(header_row),
                "targetColumn": target_column,
                "rowOffset": int(row_offset),
                "populateImages": bool(populate_images),
                "populateMSRP": bool(populate_msrp),
            },
        )

    def write_excel_generic(
        self,
        template_path: str,
        temp_dir: str,
        image_data: list[dict],
        header_row: int,
        row_offset: int,
        file_type_id: Optional[int] = None,
    ) -> Dict[str, Any]:
        return self._send_request(
            "writeExcelGeneric",
            {
                "templatePath": template_path,
                "tempDir": temp_dir,
                "imageData": image_data,
                "headerRow": int(header_row),
                "rowOffset": int(row_offset),
                "fileTypeId": file_type_id,
            },
        )


_bridge_instance: Optional[ExcelJSBridge] = None
_bridge_lock = threading.Lock()


def get_excel_bridge(logger_instance: Optional[logging.Logger] = None) -> ExcelJSBridge:
    global _bridge_instance
    with _bridge_lock:
        if _bridge_instance is None:
            _bridge_instance = ExcelJSBridge(logger=logger_instance)
        elif logger_instance is not None:
            _bridge_instance._logger = logger_instance

        _bridge_instance._ensure_process()
        return _bridge_instance


def _cleanup_bridge() -> None:
    global _bridge_instance
    with _bridge_lock:
        if _bridge_instance is not None:
            _bridge_instance.close()
            _bridge_instance = None


atexit.register(_cleanup_bridge)
