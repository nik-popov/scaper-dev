import asyncio
import logging
import os
import shutil
import time
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from typing import Optional, List, Tuple, Dict
from uuid import uuid4
import datetime
import aiohttp
import pandas as pd
import pyodbc
import requests
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from aiohttp import ClientTimeout
from aiohttp_retry import RetryClient, ExponentialRetry
from fastapi import FastAPI, BackgroundTasks
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import column_index_from_string
from PIL import Image as PILImage
from PIL import UnidentifiedImageError, ImageOps
from tldextract import tldextract
import urllib.parse
from pathlib import Path
import numpy as np
from collections import Counter
import re
import random
from functools import partial
from app_config import engine, conn_str
from email_utils import send_email, send_message_email
from s3_utils import upload_file_to_space
from openpyxl.utils.units import pixels_to_EMU, points_to_pixels
# --- Global Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
app = FastAPI()

# --- Constants ---
MAX_THREADS = int(os.environ.get('MAX_THREADS', 10))
VALID_IMAGE_TYPES = ['image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/webp', 'image/avif', 'image/tiff', 'image/x-icon']
MIN_IMAGE_SIZE = 1000  # Minimum content length in bytes
MAX_IMAGE_DIMENSION = 130  # For resizing
MIN_DIMENSION_FOR_BORDER = 10  # Minimum dimension to attempt border removal
BORDER_CROP_WIDTH = 5  # Pixels to check for border on each side
BORDER_UNIFORMITY_THRESHOLD = 0.05  # 5% of pixels must match dominant color
BORDER_COLOR_TOLERANCE = 10  # RGB tolerance for border color matching
EMU_PER_PIXEL = 9525  # EMUs per pixel at 96 DPI
def points_to_pixels(points):
    """Convert points to pixels assuming 96 DPI (1 point = 1.333 pixels)."""
    return points * 1.333
def pixels_to_EMU(pixels: float) -> int:
    """Convert pixels to English Metric Units (EMU) using 96 DPI."""
    return int(round(pixels * EMU_PER_PIXEL))
def get_user_agents(logger_instance: logging.Logger) -> List[str]:
    """
    Fetches a list of user agents from the DataProxy API.
    Returns a default list if the API call fails.
    """
    default_ua = ["Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"]
    try:
        api_url = "https://api.thedataproxy.com/v2/user-agents/?skip=0&limit=100"
        logger_instance.info(f"Fetching user agents from {api_url}")
        response = requests.get(api_url, timeout=15)
        response.raise_for_status()
        data = response.json()
        user_agents = [ua.get("user_agent") for ua in data.get("data", []) if ua.get("user_agent")]
        if not user_agents:
            logger_instance.warning("User-Agent API returned no agents. Using default.")
            return default_ua
        logger_instance.info(f"Successfully fetched {len(user_agents)} user agents.")
        return user_agents
    except requests.exceptions.RequestException as e:
        logger_instance.error(f"Failed to fetch user agents from API: {e}. Using default.")
        return default_ua

# --- General Utility Functions ---
def setup_logging(file_id: str, timestamp: str) -> logging.Logger:
    """Configure logging to save logs to a file for a specific job run."""
    log_dir = Path('jobs') / file_id / timestamp / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f'job_{file_id}_{timestamp}.log'
    
    logger_instance = logging.getLogger(f"job_{file_id}_{timestamp}")
    logger_instance.setLevel(logging.INFO)
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    if not logger_instance.handlers:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(formatter)
        
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        
        logger_instance.addHandler(file_handler)
        logger_instance.addHandler(console_handler)
        logger_instance.propagate = False
    
    logger_instance.info(f"Logging initialized for FileID: {file_id}, Timestamp: {timestamp}")
    return logger_instance

async def create_temp_dirs(unique_id: str) -> Tuple[str, str]:
    """Create temporary directories for a job."""
    base_dir = Path.cwd() / 'temp_files'
    temp_images_dir = base_dir / 'images' / unique_id
    temp_excel_dir = base_dir / 'excel' / unique_id
    
    loop = asyncio.get_running_loop()
    await loop.run_in_executor(None, partial(os.makedirs, temp_images_dir, exist_ok=True))
    await loop.run_in_executor(None, partial(os.makedirs, temp_excel_dir, exist_ok=True))
    
    return str(temp_images_dir), str(temp_excel_dir)

async def cleanup_temp_dirs(directories: List[str]):
    """Remove temporary directories."""
    loop = asyncio.get_running_loop()
    for dir_path in directories:
        await loop.run_in_executor(None, partial(shutil.rmtree, dir_path, ignore_errors=True))

# --- Database Functions ---
def get_file_type_id(file_id: int, logger_instance: logging.Logger) -> Optional[int]:
    """Retrieve the FileTypeID from the database."""
    with pyodbc.connect(conn_str) as connection:
        cursor = connection.cursor()
        query = "SELECT FileTypeID FROM utb_ImageScraperFiles WHERE ID = ?"
        cursor.execute(query, (file_id,))
        result = cursor.fetchone()
        logger_instance.info(f"FileTypeID for {file_id} is {result[0] if result else 'Not Found'}")
        return result[0] if result else None

def get_file_location_and_header(file_id: int, logger_instance: logging.Logger) -> Tuple[str, Optional[int], Optional[str]]:
    """Retrieve the source file location URL, header row index, and user email from the database."""
    with pyodbc.connect(conn_str) as connection:
        cursor = connection.cursor()
        query = "SELECT FileLocationUrl, UserHeaderIndex, UserEmail FROM utb_ImageScraperFiles WHERE ID = ?"
        cursor.execute(query, (file_id,))
        result = cursor.fetchone()
        if not result:
            raise FileNotFoundError(f"No file location found in DB for FileID {file_id}")
        file_location = result[0]
        header_row = result[1]
        user_email = result[2]
        if header_row is not None:
            try:
                header_row = int(header_row)
            except (ValueError, TypeError):
                logger_instance.warning(f"Invalid UserHeaderIndex for FileID {file_id}: {header_row}. Setting to None.")
                header_row = None
        logger_instance.info(f"FileLocationUrl for FileID {file_id}: {file_location}")
        logger_instance.info(f"UserHeaderIndex for FileID {file_id}: {header_row if header_row is not None else 'Not Found'}")
        logger_instance.info(f"UserEmail for FileID {file_id}: {user_email}")
        return file_location, header_row, user_email

def get_images_excel_db(file_id: int, logger_instance: logging.Logger) -> pd.DataFrame:
    """
    Retrieve detailed image and product data for Excel processing.
    Includes all rows, even those without image results, using LEFT JOIN.
    """
    with pyodbc.connect(conn_str) as connection:
        cursor = connection.cursor()
        cursor.execute("UPDATE utb_ImageScraperFiles SET CreateFileStartTime = GETDATE() WHERE ID = ?", (file_id,))
        connection.commit()
    
    query = """
        SELECT
            s.ExcelRowID, r.ImageUrl, r.ImageUrlThumbnail, r.SortOrder,
            s.ProductBrand AS Brand, s.ProductModel AS Style,
            s.ProductColor AS Color, s.ProductCategory AS Category,
            s.ProductMSRP AS MSRP
        FROM utb_ImageScraperFiles f
        INNER JOIN utb_ImageScraperRecords s ON s.FileID = f.ID 
        LEFT JOIN utb_ImageScraperResult r ON r.EntryID = s.EntryID AND r.SortOrder > 0
        WHERE f.ID = ?
        ORDER BY s.ExcelRowID, r.SortOrder
    """
    
    logger_instance.info(f"Executing query to fetch all records for FileID: {file_id}")
    return pd.read_sql_query(query, engine, params=[(file_id,)])

def update_file_location_complete(file_id: int, file_location: str, logger_instance: logging.Logger):
    """Update the completed file location URL in the database."""
    with pyodbc.connect(conn_str) as connection:
        cursor = connection.cursor()
        query = "UPDATE utb_ImageScraperFiles SET FileLocationURLComplete = ? WHERE ID = ?"
        cursor.execute(query, (file_location, file_id))
        connection.commit()
        logger_instance.info(f"Updated completed file location for FileID {file_id}")

# --- Image Downloading Functions ---
async def validate_image_response(response, url: str, logger_instance: logging.Logger) -> Tuple[bool, Optional[bytes]]:
    if response.status != 200:
        logger_instance.warning(f"Invalid status code {response.status} for URL: {url}")
        return False, None
    if not any(response.headers.get('Content-Type', '').lower().startswith(it) for it in VALID_IMAGE_TYPES):
        logger_instance.warning(f"Invalid Content-Type for URL: {url}")
        return False, None
    if int(response.headers.get('Content-Length', 0)) < MIN_IMAGE_SIZE:
        logger_instance.warning(f"Content too small for URL: {url}")
        return False, None
    return True, await response.read()

async def image_download(semaphore, item: Dict, save_path: str, session, logger_instance: logging.Logger, user_agents: List[str]):
    async with semaphore:
        row_id = item['ExcelRowID']
        image_name = str(row_id)
        loop = asyncio.get_running_loop()

        for i, (url, thumb_url, sort_order) in enumerate(item['image_options']):
            
            async def attempt_download(download_url: str, is_thumb: bool) -> bool:
                if not download_url: return False
                try:
                    headers = {"User-Agent": random.choice(user_agents)}
                    async with session.get(download_url, headers=headers) as response:
                        is_valid, data = await validate_image_response(response, download_url, logger_instance)
                        if not is_valid: return False
                        
                        final_path = os.path.join(save_path, f"{image_name}.png")

                        def save_image_sync():
                            with PILImage.open(BytesIO(data)) as img:
                                # Ensure proper mode conversion during save
                                if img.mode == 'RGBA':
                                    background = PILImage.new('RGB', img.size, (255, 255, 255))
                                    background.paste(img, mask=img.split()[3])
                                    img = background
                                elif img.mode != 'RGB':
                                    img = img.convert('RGB')
                                img.save(final_path, 'PNG', compress_level=6)
                        
                        await loop.run_in_executor(None, save_image_sync)
                                
                        logger_instance.info(f"SUCCESS (SortOrder {sort_order}) for Row {row_id} from {'thumbnail' if is_thumb else 'main'} URL.")
                        return True
                except Exception as e:
                    logger_instance.warning(f"Download attempt {i+1} failed for Row {row_id} ({'thumb' if is_thumb else 'main'}): {e}")
                    return False

            if await attempt_download(url, is_thumb=False):
                return True
            if await attempt_download(thumb_url, is_thumb=True):
                return True
        
        logger_instance.error(f"All download attempts FAILED for Row {row_id}. No more images to try.")
        return False

async def download_all_images(data: List[Dict], save_path: str, logger_instance: logging.Logger, user_agents: List[str]):
    """Download all images concurrently, trying fallbacks for each item."""
    domains = []
    for item in data:
        if item.get('image_options'):
            first_url = item['image_options'][0][0]
            if first_url:
                domains.append(tldextract.extract(first_url).top_domain_under_public_suffix)
    
    pool_size = min(500, max(10, len(Counter(domains)) * 2))
    timeout_config = ClientTimeout(total=60)
    retry_options = ExponentialRetry(attempts=3, start_timeout=3)
    
    async with RetryClient(raise_for_status=False, retry_options=retry_options, timeout=timeout_config) as session:
        semaphore = asyncio.Semaphore(pool_size)
        tasks = [image_download(semaphore, item, save_path, session, logger_instance, user_agents) for item in data]
        results = await asyncio.gather(*tasks)
        
        failed_count = len([res for res in results if not res])
        if failed_count > 0:
            logger_instance.warning(f"{failed_count} product rows failed to download any image.")

# --- Image & Excel Processing Functions ---
def resize_image(image_path: str, logger_instance: logging.Logger) -> bool:
    try:
        with PILImage.open(image_path) as img:
            img.load()  # Ensure image is fully loaded

            # Resize in memory (upscale or downscale)
            width, height = img.size
            ratio = min(MAX_IMAGE_DIMENSION / width, MAX_IMAGE_DIMENSION / height)
            new_w = int(width * ratio)
            new_h = int(height * ratio)
            
            if (new_w, new_h) != (width, height):
                img = img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                logger_instance.info(f"Resized image {image_path} from {width}x{height} to {new_w}x{new_h} (max dim {MAX_IMAGE_DIMENSION})")

            # Save once at the end
            img.save(image_path, 'PNG', compress_level=6, quality=95)
            return True
    except Exception as e:
        logger_instance.error(f"Error resizing image {image_path}: {e}", exc_info=True)
        return False

def get_last_non_empty_row(ws, column: str, header_row: int, logger_instance: logging.Logger) -> int:
    """Find the last non-empty row in the specified column, starting after header_row."""
    last_row = header_row
    for row in ws[f"{column}{header_row + 1}:{column}{ws.max_row}"]:
        if row[0].value is not None and str(row[0].value).strip():
            last_row = max(last_row, row[0].row)
    logger_instance.info(f"Last non-empty row in column {column}: {last_row}")
    return last_row

def verify_and_process_image(image_path: str, logger_instance: logging.Logger) -> bool:
    """Verify, crop, and process image in memory to avoid corruption."""
    try:
        # Verify image integrity
        with PILImage.open(image_path) as temp_img:
            temp_img.verify()
        if os.path.getsize(image_path) < MIN_IMAGE_SIZE:
            logger_instance.warning(f"File too small: {image_path}")
            return False

        # Load image for processing
        img = PILImage.open(image_path)
        img.load()

        # Process image (cropping and border removal)
        img = process_image_remove_lines(image_path, img, logger_instance)

        # Resize to fit max dimension (upscale or downscale)
        w, h = img.size
        ratio = min(MAX_IMAGE_DIMENSION / w, MAX_IMAGE_DIMENSION / h)
        new_w = int(w * ratio)
        new_h = int(h * ratio)
        
        if (new_w, new_h) != (w, h):
            img = img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
            logger_instance.info(f"Resized image {image_path} from {w}x{h} to {new_w}x{new_h} (max dim {MAX_IMAGE_DIMENSION})")

        # Save processed image once
        img.save(image_path, 'PNG', compress_level=6, quality=95)
        logger_instance.info(f"Image verified and processed: {image_path}")
        return True
    except Exception as e:
        logger_instance.error(f"Image verification failed for {image_path}: {e}", exc_info=True)
        return False

def process_image_remove_lines(image_path: str, img: PILImage.Image, logger_instance: logging.Logger) -> PILImage.Image:
    """Process image in memory to remove uniform lines and borders, preventing smearing."""
    try:
        img = ImageOps.exif_transpose(img.copy())
        arr = np.array(img)
        original_height, original_width, _ = arr.shape

        # Detect non-uniform horizontal rows
        valid_row_indices = [y for y in range(original_height) if np.any(np.max(arr[y], axis=0) - np.min(arr[y], axis=0) >= 15)]
        row_cropped = False
        if valid_row_indices:
            rows_to_keep = get_valid_indices_with_padding(valid_row_indices, original_height - 1, pad=5)
            row_cropped = len(rows_to_keep) != original_height
            arr = arr[rows_to_keep]
            logger_instance.info(f"Kept {len(rows_to_keep)} of {original_height} rows for {image_path}")
        else:
            logger_instance.info(f"No non-uniform rows found in {image_path}")

        # Detect non-uniform vertical columns
        arr_t = np.transpose(arr, (1, 0, 2))
        new_width = arr_t.shape[0]
        valid_col_indices = [x for x in range(new_width) if np.any(np.max(arr_t[x], axis=0) - np.min(arr_t[x], axis=0) >= 15)]
        col_cropped = False
        if valid_col_indices:
            cols_to_keep = get_valid_indices_with_padding(valid_col_indices, new_width - 1, pad=5)
            col_cropped = len(cols_to_keep) != new_width
            arr_t = arr_t[cols_to_keep]
            arr = np.transpose(arr_t, (1, 0, 2))
            logger_instance.info(f"Kept {len(cols_to_keep)} of {new_width} columns for {image_path}")
        else:
            logger_instance.info(f"No non-uniform columns found in {image_path}")

        # Check if cropped image is too small
        cleaned_height, cleaned_width, _ = arr.shape
        min_dim = 20
        if (row_cropped or col_cropped) and (cleaned_height < min_dim or cleaned_width < min_dim):
            logger_instance.warning(f"Cropped image too small ({cleaned_height}x{cleaned_width}) for {image_path}. Using original.")
            return img

        # Border removal (only if dimensions are sufficient)
        pixels = arr
        if cleaned_height >= MIN_DIMENSION_FOR_BORDER and cleaned_width >= MIN_DIMENSION_FOR_BORDER:
            try:
                # Collect border pixels from outer 5 pixels, avoiding corner overlaps
                top_border = pixels[:BORDER_CROP_WIDTH, BORDER_CROP_WIDTH:-BORDER_CROP_WIDTH, :]
                bottom_border = pixels[-BORDER_CROP_WIDTH:, BORDER_CROP_WIDTH:-BORDER_CROP_WIDTH, :]
                left_border = pixels[BORDER_CROP_WIDTH:-BORDER_CROP_WIDTH, :BORDER_CROP_WIDTH, :]
                right_border = pixels[BORDER_CROP_WIDTH:-BORDER_CROP_WIDTH, -BORDER_CROP_WIDTH:, :]
                
                border_pixels = np.vstack([
                    top_border.reshape(-1, 3),
                    bottom_border.reshape(-1, 3),
                    left_border.reshape(-1, 3),
                    right_border.reshape(-1, 3)
                ])
                
                # Get dominant border color
                colors, counts = np.unique(border_pixels, axis=0, return_counts=True)
                total_border_pixels = len(border_pixels)
                dominant_color = colors[np.argmax(counts)]
                dominant_fraction = counts.max() / total_border_pixels
                
                # Check if border is light and sufficiently uniform
                if np.sum(dominant_color) > 750 and dominant_fraction > BORDER_UNIFORMITY_THRESHOLD:
                    # Create mask for border areas only
                    mask = np.zeros((cleaned_height, cleaned_width), dtype=bool)
                    mask[:BORDER_CROP_WIDTH, :] = np.all(np.abs(pixels[:BORDER_CROP_WIDTH, :, :] - dominant_color) <= BORDER_COLOR_TOLERANCE, axis=2)
                    mask[-BORDER_CROP_WIDTH:, :] = np.all(np.abs(pixels[-BORDER_CROP_WIDTH:, :, :] - dominant_color) <= BORDER_COLOR_TOLERANCE, axis=2)
                    mask[:, :BORDER_CROP_WIDTH] = np.all(np.abs(pixels[:, :BORDER_CROP_WIDTH, :] - dominant_color) <= BORDER_COLOR_TOLERANCE, axis=2)
                    mask[:, -BORDER_CROP_WIDTH:] = np.all(np.abs(pixels[:, -BORDER_CROP_WIDTH:, :] - dominant_color) <= BORDER_COLOR_TOLERANCE, axis=2)
                    
                    # Avoid modifying inner region
                    mask[BORDER_CROP_WIDTH:-BORDER_CROP_WIDTH, BORDER_CROP_WIDTH:-BORDER_CROP_WIDTH] = False
                    
                    # Replace border pixels with white
                    pixels[mask] = np.array([255, 255, 255])
                    arr = pixels
                    logger_instance.info(f"Applied border removal for {image_path} (dominant color {dominant_color}, fraction {dominant_fraction:.2f})")
                else:
                    logger_instance.info(f"No light uniform border detected for {image_path} (fraction {dominant_fraction:.2f}, color sum {np.sum(dominant_color)})")
            except Exception as e:
                logger_instance.warning(f"Border removal failed for {image_path}: {e}. Using cropped image.")
        else:
            logger_instance.info(f"Image too small for border removal ({cleaned_height}x{cleaned_width}) for {image_path}")

        return PILImage.fromarray(arr.astype(np.uint8))
    except Exception as e:
        logger_instance.error(f"Error processing image {image_path} to remove lines: {e}", exc_info=True)
        return img

def get_valid_indices_with_padding(valid_indices, max_index, pad=5):
    """Identify indices with padding for non-uniform rows/columns."""
    padded = set()
    for idx in valid_indices:
        start = max(0, idx - pad)
        end = min(max_index, idx + pad + 1)
        padded.update(range(start, end))
    return sorted(padded)

def write_excel_distro(local_filename: str, temp_dir: str, image_data: List[Dict], header_row: int, logger_instance: logging.Logger, row_offset: int = 0):
    try:
        header_row = int(header_row)
        if header_row < 0:
            logger_instance.error(f"Invalid header_row {header_row}. Setting to 0.")
            header_row = 0
        if row_offset < 0:
            logger_instance.warning(f"Negative row_offset {row_offset} provided. Ensure this is intentional.")

        wb = load_workbook(local_filename)
        ws = wb.active
        image_map = {}
        for path_obj in Path(temp_dir).iterdir():
            if not path_obj.is_file():
                continue
            stem = path_obj.stem
            if stem.isdigit():
                image_map[int(stem)] = str(path_obj)

        row_dim = ws.row_dimensions.get(header_row + 1)
        if row_dim is not None and getattr(row_dim, "height", None) is not None:
            DEFAULT_ROW_HEIGHT_POINTS = row_dim.height
        else:
            DEFAULT_ROW_HEIGHT_POINTS = 12.75
        logger_instance.info(f"Using template row height: {DEFAULT_ROW_HEIGHT_POINTS} points from row {header_row + 1}")

        row_data_map = {item['ExcelRowID']: item for item in image_data}
        last_non_empty_row = get_last_non_empty_row(ws, column='B', header_row=header_row, logger_instance=logger_instance)

        data_row_ids = sorted(row_data_map.keys())
        if data_row_ids:
            min_row_id = data_row_ids[0]
            max_row_id = data_row_ids[-1]
        else:
            min_row_id = 1
            relative_last_row = last_non_empty_row - (header_row + 1)
            relative_last_row = max(1, relative_last_row)
            max_row_id = relative_last_row
            logger_instance.warning(
                f"No data in image_data, deriving row range 1-{max_row_id} from template last row {last_non_empty_row}"
            )

        base_row = header_row + row_offset + 2
        if base_row < 1:
            logger_instance.error(f"Computed base_row {base_row} is less than 1. Aborting.")
            raise ValueError("Invalid row range: base_row is less than 1")

        total_rows_needed = max(0, max_row_id - min_row_id + 1)
        max_needed_row = base_row + max(0, total_rows_needed - 1)
        max_needed_row = max(max_needed_row, last_non_empty_row)
        logger_instance.info(
            f"Row mapping summary -> base_row: {base_row}, min_row_id: {min_row_id}, max_row_id: {max_row_id}, "
            f"total_rows_needed: {total_rows_needed}, max_needed_row: {max_needed_row}"
        )

        if ws.max_row < max_needed_row:
            logger_instance.info(f"Appending {max_needed_row - ws.max_row} rows to worksheet")
            for row_num in range(ws.max_row + 1, max_needed_row + 1):
                ws.append([''] * ws.max_column)
                ws.row_dimensions[row_num].height = DEFAULT_ROW_HEIGHT_POINTS

        CELL_WIDTH_POINTS = 100
        CELL_HEIGHT_POINTS = max(DEFAULT_ROW_HEIGHT_POINTS, 150)
        PADDING_POINTS = 5

        for row_id in range(min_row_id, max_row_id + 1):
            row_position = row_id - min_row_id
            row_num = base_row + row_position

            ws.row_dimensions[row_num].height = CELL_HEIGHT_POINTS

            image_path = image_map.get(row_id)
            if image_path:
                if verify_and_process_image(image_path, logger_instance):
                    img = Image(image_path)
                    img_height_pixels = getattr(img, 'height', 0)
                    img_width_pixels = getattr(img, 'width', 0)
                    img_height_points = img_height_pixels * 72 / 96
                    img_width_points = img_width_pixels / 7

                    required_width = img_width_points + PADDING_POINTS * 2
                    current_width = ws.column_dimensions['A'].width or 8.43
                    ws.column_dimensions['A'].width = min(CELL_WIDTH_POINTS, max(current_width, required_width))

                    cell_width_pixels = ws.column_dimensions['A'].width * 7
                    cell_height_pixels = points_to_pixels(CELL_HEIGHT_POINTS)
                    x_offset_pixels = max(0, (cell_width_pixels - img_width_pixels) / 2)
                    y_offset_pixels = max(0, (cell_height_pixels - img_height_pixels) / 2)

                    width_emu = pixels_to_EMU(img_width_pixels)
                    height_emu = pixels_to_EMU(img_height_pixels)
                    marker = AnchorMarker(
                        col=0,
                        colOff=pixels_to_EMU(x_offset_pixels),
                        row=max(0, row_num - 1),
                        rowOff=pixels_to_EMU(y_offset_pixels)
                    )
                    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(width_emu, height_emu))
                    ws.add_image(img)
                    logger_instance.info(
                        f"Added image for Row {row_id} at Excel row {row_num}, "
                        f"height={CELL_HEIGHT_POINTS} points, width={ws.column_dimensions['A'].width:.2f} points"
                    )
                else:
                    logger_instance.warning(f"Image processing failed for Row {row_id}, writing metadata only")
            else:
                logger_instance.info(f"No image found for Row {row_id}, writing metadata only")

            item = row_data_map.get(row_id)
            if item:
                if row_num > header_row + 1:
                    ws[f"B{row_num}"] = item.get('Brand', '')
                    ws[f"D{row_num}"] = item.get('Style', '')
                    ws[f"E{row_num}"] = item.get('Color', '')
                    ws[f"H{row_num}"] = item.get('Category', '')
                    logger_instance.info(f"Wrote metadata for Row {row_id} at Excel row {row_num}")
                else:
                    logger_instance.warning(f"Skipping metadata write for row {row_num} as it is the header row")
            else:
                if row_num > header_row + 1:
                    ws[f"B{row_num}"] = ''
                    ws[f"D{row_num}"] = ''
                    ws[f"E{row_num}"] = ''
                    ws[f"H{row_num}"] = ''
                    logger_instance.info(f"Filled missing row {row_num} (ExcelRowID {row_id}) with empty metadata")
                else:
                    logger_instance.info(f"Skipping row {row_num} (ExcelRowID {row_id}) as it is the header row")

        if ws.max_row > max_needed_row:
            logger_instance.info(f"Deleting {ws.max_row - max_needed_row} rows after row {max_needed_row}")
            ws.delete_rows(max_needed_row + 1, ws.max_row - max_needed_row)

        logger_instance.info("Setting worksheet view to A1")
        ws.sheet_view.topLeftCell = 'A1'
        wb.save(local_filename)
        logger_instance.info(f"Excel file saved: {local_filename}")
    except Exception as e:
        logger_instance.error(f"Error writing to DISTRO Excel file: {e}", exc_info=True)
        raise


def find_header_row_index(excel_file: str, logger_instance: logging.Logger) -> Optional[int]:
    try:
        wb = load_workbook(excel_file, read_only=True)
        ws = wb.active
        keywords = {'id', 'name', 'product', 'brand', 'sku', 'category', 'color', 'price', 'description'}
        best_row, best_score = None, 0
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            if not row:
                continue
            cells = [str(c).strip().lower() for c in row if isinstance(c, str)]
            score = len(cells) + len(set(cells)) + sum(any(k in c for k in keywords) for c in cells)
            if score > best_score:
                best_score, best_row = score, i
        if best_row and best_score > 3:
            return best_row - 1  # Convert to 0-based indexing
        return None
    except Exception as e:
        logger_instance.error(f"Could not find header row: {e}")
        return None

def write_excel_msrp(local_filename: str, temp_dir: str, image_data: List[Dict], header_row: int, target_column: str, row_offset: int, logger_instance: logging.Logger, populate_images: bool = True, populate_msrp: bool = True):
    try:
        wb = load_workbook(local_filename)
        ws = wb.active
        image_map = {int(Path(f).stem): f for f in os.listdir(temp_dir) if Path(f).stem.isdigit()}
        if populate_msrp and not re.match(r'^[A-Z]+$', target_column):
            raise ValueError(f"Invalid target_column: {target_column}. Must be a valid Excel column letter (e.g., 'A', 'B', 'AA').")

        header_row = int(header_row)
        if header_row < 0:
            logger_instance.error(f"Invalid header_row {header_row}. Setting to 0.")
            header_row = 0
        if row_offset < 0:
            logger_instance.warning(f"Negative row_offset {row_offset} provided. Ensure this is intentional.")

        row_dim = ws.row_dimensions.get(header_row + 1)
        DEFAULT_ROW_HEIGHT_POINTS = row_dim.height if row_dim and row_dim.height is not None else 12.75
        logger_instance.info(f"Using template row height: {DEFAULT_ROW_HEIGHT_POINTS} points from row {header_row + 1}")

        row_data_map = {item['ExcelRowID']: item for item in image_data}
        last_non_empty_row = get_last_non_empty_row(ws, column='B', header_row=header_row, logger_instance=logger_instance)

        data_row_ids = sorted(row_data_map.keys())
        if data_row_ids:
            min_row_id = data_row_ids[0]
            max_row_id = data_row_ids[-1]
        else:
            min_row_id = 1
            relative_last_row = last_non_empty_row - (header_row + 1)
            relative_last_row = max(1, relative_last_row)
            max_row_id = relative_last_row
            logger_instance.warning(
                f"No data in image_data, deriving row range 1-{max_row_id} from template last row {last_non_empty_row}"
            )

        base_row = header_row + row_offset + 2
        if base_row < 1:
            logger_instance.error(f"Computed base_row {base_row} is less than 1. Aborting.")
            raise ValueError("Invalid row range: base_row is less than 1")

        total_rows_needed = max(0, max_row_id - min_row_id + 1)
        max_needed_row = base_row + max(0, total_rows_needed - 1)
        max_needed_row = max(max_needed_row, last_non_empty_row)
        logger_instance.info(
            f"Row mapping summary -> base_row: {base_row}, min_row_id: {min_row_id}, max_row_id: {max_row_id}, "
            f"total_rows_needed: {total_rows_needed}, max_needed_row: {max_needed_row}"
        )

        if ws.max_row < max_needed_row:
            logger_instance.info(f"Appending {max_needed_row - ws.max_row} rows to worksheet")
            for row_num in range(ws.max_row + 1, max_needed_row + 1):
                ws.append([''] * ws.max_column)
                ws.row_dimensions[row_num].height = DEFAULT_ROW_HEIGHT_POINTS

        for row_id in range(min_row_id, max_row_id + 1):
            row_position = row_id - min_row_id
            row_num = base_row + row_position

            ws.row_dimensions[row_num].height = DEFAULT_ROW_HEIGHT_POINTS

            if row_id in row_data_map:
                item = row_data_map[row_id]
                if populate_images:
                    if row_id in image_map:
                        image_path = os.path.join(temp_dir, image_map[row_id])
                        if verify_and_process_image(image_path, logger_instance):
                            img = Image(image_path)
                            
                            img_height_pixels = img.height if hasattr(img, 'height') else 0
                            img_width_pixels = img.width if hasattr(img, 'width') else 0
                            img_height_points = img_height_pixels * 72 / 96
                            
                            new_height = max(DEFAULT_ROW_HEIGHT_POINTS, img_height_points)
                            ws.row_dimensions[row_num].height = new_height
                            
                            # Calculate dimensions for centering
                            col_width_chars = ws.column_dimensions['A'].width
                            if col_width_chars is None:
                                col_width_chars = 8.43
                            cell_width_pixels = col_width_chars * 7
                            cell_height_pixels = points_to_pixels(new_height)
                            
                            x_offset_pixels = max(0, (cell_width_pixels - img_width_pixels) / 2)
                            y_offset_pixels = max(0, (cell_height_pixels - img_height_pixels) / 2)
                            
                            width_emu = pixels_to_EMU(img_width_pixels)
                            height_emu = pixels_to_EMU(img_height_pixels)
                            
                            marker = AnchorMarker(
                                col=0, # Column A
                                colOff=pixels_to_EMU(x_offset_pixels),
                                row=row_num - 1,
                                rowOff=pixels_to_EMU(y_offset_pixels)
                            )
                            img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(width_emu, height_emu))
                            ws.add_image(img)
                            
                            logger_instance.info(f"Added image for Row {row_id} at Excel row {row_num}, height set to {ws.row_dimensions[row_num].height} points")
                        else:
                            logger_instance.warning(f"Image processing failed for Row {row_id}, writing MSRP only")
                    else:
                        logger_instance.info(f"No image found for Row {row_id}, writing MSRP only")

                if populate_msrp:
                    msrp_value = item.get('MSRP', '')
                    ws.cell(row=row_num, column=column_index_from_string(target_column)).value = msrp_value
                    logger_instance.info(f"Wrote MSRP '{msrp_value}' for Row {row_id} at {target_column}{row_num}")
            else:
                if populate_msrp:
                    ws.cell(row=row_num, column=column_index_from_string(target_column)).value = ''
                    logger_instance.info(f"Filled missing row {row_num} (ExcelRowID {row_id}) with empty MSRP")

        if ws.max_row > max_needed_row:
            logger_instance.info(f"Deleting {ws.max_row - max_needed_row} rows after row {max_needed_row}")
            ws.delete_rows(max_needed_row + 1, ws.max_row - max_needed_row)

        logger_instance.info("Setting worksheet view to A1")
        ws.sheet_view.topLeftCell = 'A1'
        wb.save(local_filename)
        logger_instance.info(f"MSRP Excel file saved: {local_filename}")
    except Exception as e:
        logger_instance.error(f"Error writing to MSRP Excel file: {e}", exc_info=True)
        raise

def write_excel_generic(local_filename: str, temp_dir: str, image_data: List[Dict], header_row: int, row_offset: int, logger_instance: logging.Logger, file_type_id: Optional[int] = None):
    try:
        wb = load_workbook(local_filename)
        ws = wb.active

        # Clear existing images for FileTypeID 10
        if file_type_id == 10:
            if hasattr(ws, '_images'):
                ws._images = []
                logger_instance.info("Cleared existing images for FileTypeID 10")

        temp_path = Path(temp_dir)
        image_map = {}
        for path_obj in temp_path.iterdir():
            if not path_obj.is_file():
                continue
            stem = path_obj.stem
            if stem.isdigit():
                image_map[int(stem)] = path_obj

        row_ids_from_data = set()
        for item in image_data:
            value = item.get('ExcelRowID')
            if value is None:
                continue
            try:
                row_ids_from_data.add(int(value))
            except (ValueError, TypeError):
                logger_instance.warning(f"Skipping invalid ExcelRowID '{value}' in generic writer")

        all_row_ids = sorted(row_ids_from_data.union(image_map.keys()))

        if not all_row_ids:
            logger_instance.info("No rows found to write for generic template.")
            return

        min_row_id = all_row_ids[0]
        max_row_id = all_row_ids[-1]

        template_row_dim = ws.row_dimensions.get(header_row + 1)
        default_row_height = template_row_dim.height if template_row_dim and template_row_dim.height is not None else 12.75

        header_excel_row = max(header_row, 0) + 1
        use_absolute_row_ids = min_row_id > header_excel_row
        if use_absolute_row_ids:
            logger_instance.info(
                f"Detected absolute ExcelRowID values (min_row_id={min_row_id}, header_excel_row={header_excel_row})."
            )
        else:
            logger_instance.info(
                f"Using relative ExcelRowID mapping (min_row_id={min_row_id}, header_excel_row={header_excel_row})."
            )

        row_num_map: Dict[int, int] = {}
        total_columns = max(ws.max_column, 1)
        relative_base_row = header_row + row_offset + 2
        for row_id in all_row_ids:
            if use_absolute_row_ids:
                row_num = row_id + row_offset
            else:
                row_num = relative_base_row + (row_id - min_row_id)

            if row_num < 1:
                logger_instance.error(
                    f"Computed row {row_num} is invalid for row_id={row_id}, header_row={header_row}, row_offset={row_offset}. Skipping."
                )
                continue
            row_num_map[row_id] = row_num

        if not row_num_map:
            logger_instance.warning("No valid row mappings computed for generic template. Nothing to write.")
            return

        max_needed_row = max(row_num_map.values())
        if ws.max_row < max_needed_row:
            for _ in range(ws.max_row + 1, max_needed_row + 1):
                ws.append([''] * total_columns)

        for row_id in all_row_ids:
            row_num = row_num_map.get(row_id)
            if row_num is None:
                continue

            image_path_obj = image_map.get(row_id)
            if not image_path_obj:
                continue

            image_path = str(image_path_obj)
            if verify_and_process_image(image_path, logger_instance):
                img = Image(image_path)
                
                img_height_pixels = getattr(img, 'height', 0)
                img_width_pixels = getattr(img, 'width', 0)
                img_height_points = img_height_pixels * 72 / 96
                current_height = ws.row_dimensions[row_num].height
                new_height = max(img_height_points, current_height or default_row_height)
                ws.row_dimensions[row_num].height = new_height
                
                # Calculate dimensions for centering
                col_width_chars = ws.column_dimensions['A'].width
                if col_width_chars is None:
                    col_width_chars = 8.43
                cell_width_pixels = col_width_chars * 7
                cell_height_pixels = points_to_pixels(new_height)
                
                x_offset_pixels = max(0, (cell_width_pixels - img_width_pixels) / 2)
                y_offset_pixels = max(0, (cell_height_pixels - img_height_pixels) / 2)
                
                width_emu = pixels_to_EMU(img_width_pixels)
                height_emu = pixels_to_EMU(img_height_pixels)
                
                marker = AnchorMarker(
                    col=0, # Column A
                    colOff=pixels_to_EMU(x_offset_pixels),
                    row=row_num - 1,
                    rowOff=pixels_to_EMU(y_offset_pixels)
                )
                img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(width_emu, height_emu))
                ws.add_image(img)

                logger_instance.info(
                    f"Added image for ExcelRowID {row_id} at Excel row {row_num} (absolute_mapping={use_absolute_row_ids})"
                )
            else:
                logger_instance.warning(f"Image verification failed for ExcelRowID {row_id} at path {image_path}")

        logger_instance.info("Setting worksheet view to A1")
        ws.sheet_view.topLeftCell = 'A1'
        wb.save(local_filename)
    except Exception as e:
        logger_instance.error(f"Error writing to generic Excel file: {e}", exc_info=True)
        raise

# --- Main Background Task ---
async def generate_download_file(file_id: str, row_offset: int = 0):
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    logger_instance = setup_logging(file_id, timestamp)
    temp_images_dir, temp_excel_dir = "", ""
    file_id_int = int(file_id)

    try:
        user_agents = get_user_agents(logger_instance)
        temp_images_dir, temp_excel_dir = await create_temp_dirs(file_id)
        
        logger_instance.info("Fetching all image data from database...")
        images_df = get_images_excel_db(file_id_int, logger_instance)
        
        if images_df.empty:
            logger_instance.warning(f"No image data found for FileID {file_id}. Fetching all records for metadata.")
            query = """
                SELECT
                    s.ExcelRowID,
                    s.ProductBrand AS Brand, s.ProductModel AS Style,
                    s.ProductColor AS Color, s.ProductCategory AS Category
                FROM utb_ImageScraperFiles f
                INNER JOIN utb_ImageScraperRecords s ON s.FileID = f.ID 
                WHERE f.ID = ?
                ORDER BY s.ExcelRowID
            """
            images_df = pd.read_sql_query(query, engine, params=[(file_id,)])
            if images_df.empty:
                logger_instance.error(f"No records found for FileID {file_id}. The job cannot proceed.")
                return

        logger_instance.info(f"ExcelRowID values: {list(images_df['ExcelRowID'])}")
        logger_instance.info("Grouping data by product to prepare for download attempts...")
        grouped_data = []
        for _, group in images_df.groupby('ExcelRowID'):
            first_row = group.iloc[0]
            image_options = list(zip(group['ImageUrl'], group['ImageUrlThumbnail'], group['SortOrder'])) if 'ImageUrl' in group.columns else []
            
            grouped_data.append({
                'ExcelRowID': int(first_row['ExcelRowID']),
                'Brand': first_row['Brand'],
                'Style': first_row['Style'],
                'Color': first_row['Color'],
                'Category': first_row['Category'],
                'image_options': image_options
            })
        
        logger_instance.info(f"Data prepared for {len(grouped_data)} unique products. Starting image downloads.")
        if any(item['image_options'] for item in grouped_data):
            await download_all_images(grouped_data, temp_images_dir, logger_instance, user_agents)
        else:
            logger_instance.info("No images to download, proceeding with metadata only.")
        
        file_type_id = get_file_type_id(file_id_int, logger_instance)
        FILE_TYPE_DISTRO = 3

        file_url, header_row_from_db, user_email = get_file_location_and_header(file_id_int, logger_instance)
        original_file_name = os.path.basename(urllib.parse.unquote(file_url))

        if file_type_id == FILE_TYPE_DISTRO:
            logger_instance.info("Starting DISTRO file generation.")
            template_url = "https://iconluxury.shop/documents/public_ICON_DISTRO_USD_20250617.xlsx"
            file_name = os.path.basename(urllib.parse.unquote(template_url))
            local_filename = os.path.join(temp_excel_dir, file_name)
            header_row = 4  # Hardcoded for DISTRO (0-based)

            if row_offset:
                logger_instance.info(
                    f"Ignoring provided row_offset={row_offset} for DISTRO template; using fixed layout."
                )
            template_row_offset = 0

            res = requests.get(template_url, timeout=60)
            res.raise_for_status()
            with open(local_filename, "wb") as f:
                f.write(res.content)
            write_excel_distro(local_filename, temp_images_dir, grouped_data, header_row, logger_instance, template_row_offset)
        else:
            logger_instance.info("Starting GENERIC file generation.")
            file_name = os.path.basename(urllib.parse.unquote(file_url))
            local_filename = os.path.join(temp_excel_dir, file_name)

            res = requests.get(file_url, timeout=60)
            res.raise_for_status()
            with open(local_filename, "wb") as f:
                f.write(res.content)

            inferred_header_row = find_header_row_index(local_filename, logger_instance)
            if header_row_from_db is not None:
                try:
                    db_header = int(header_row_from_db)
                except (TypeError, ValueError):
                    logger_instance.warning(f"Invalid header_row_from_db '{header_row_from_db}'. Falling back to inferred value.")
                    db_header = None

                if db_header is not None:
                    candidates = {max(0, db_header - 1), max(0, db_header)}
                    if inferred_header_row is not None:
                        header_row_value = min(candidates, key=lambda c: abs(c - inferred_header_row))
                    else:
                        header_row_value = min(candidates)
                else:
                    header_row_value = inferred_header_row if inferred_header_row is not None else 0
            else:
                header_row_value = inferred_header_row if inferred_header_row is not None else 0

            logger_instance.info(f"Using header_row_value={header_row_value} for generic template.")
            write_excel_generic(local_filename, temp_images_dir, grouped_data, header_row_value, row_offset, logger_instance, file_type_id=file_type_id)
            
        processed_file_name = f"{Path(original_file_name).stem}_ptype-{file_type_id}_{timestamp}.xlsx"
        public_url = await upload_file_to_space(local_filename, save_as=f"processed_files/{processed_file_name}", file_id=file_id_int, is_public=True)
        update_file_location_complete(file_id_int, public_url, logger_instance)
        await send_email(
            to_emails=user_email or '',
            subject=f'File Processed: {file_name}',
            file_path=local_filename,
            job_id=file_id
        )
        
        logger_instance.info(f"Successfully completed job for FileID {file_id}.")

    except Exception as e:
        logger_instance.error(f"FATAL ERROR for FileID {file_id}: {e}", exc_info=True)
    finally:
        if temp_images_dir and temp_excel_dir:
            await cleanup_temp_dirs([temp_images_dir, temp_excel_dir])

async def generate_msrp_excel(file_id: str, target_column: str, row_offset: int = 0):
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    logger_instance = setup_logging(file_id, timestamp)
    temp_images_dir, temp_excel_dir = "", ""
    file_id_int = int(file_id)

    try:
        user_agents = get_user_agents(logger_instance)
        temp_images_dir, temp_excel_dir = await create_temp_dirs(file_id)
        
        # Determine what to populate based on FileTypeID
        file_type_id = get_file_type_id(file_id_int, logger_instance)
        populate_images = True
        populate_msrp = True
        
        if file_type_id == 8:
            populate_msrp = False
            logger_instance.info(f"FileTypeID is 8: Populating IMAGES ONLY.")
        elif file_type_id == 9:
            populate_images = False
            logger_instance.info(f"FileTypeID is 9: Populating MSRP ONLY.")

        logger_instance.info("Fetching all image and MSRP data from database...")
        images_df = get_images_excel_db(file_id_int, logger_instance)
        
        if images_df.empty:
            logger_instance.warning(f"No data found for FileID {file_id}. Fetching all records for MSRP.")
            query = """
                SELECT
                    s.ExcelRowID,
                    s.ProductMSRP AS MSRP
                FROM utb_ImageScraperFiles f
                INNER JOIN utb_ImageScraperRecords s ON s.FileID = f.ID 
                WHERE f.ID = ?
                ORDER BY s.ExcelRowID
            """
            images_df = pd.read_sql_query(query, engine, params=[(file_id,)])
            if images_df.empty:
                logger_instance.error(f"No records found for FileID {file_id}. The job cannot proceed.")
                return

        logger_instance.info("Grouping data by product to prepare for download attempts...")
        grouped_data = []
        for _, group in images_df.groupby('ExcelRowID'):
            first_row = group.iloc[0]
            image_options = list(zip(group['ImageUrl'], group['ImageUrlThumbnail'], group['SortOrder'])) if 'ImageUrl' in group.columns else []
            
            grouped_data.append({
                'ExcelRowID': int(first_row['ExcelRowID']),
                'MSRP': first_row.get('MSRP', ''),
                'image_options': image_options
            })
        
        logger_instance.info(f"Data prepared for {len(grouped_data)} unique products. Starting image downloads.")
        
        if populate_images:
            if any(item['image_options'] for item in grouped_data):
                await download_all_images(grouped_data, temp_images_dir, logger_instance, user_agents)
            else:
                logger_instance.info("No images to download.")
        else:
            logger_instance.info("Skipping image download as populate_images is False.")
        
        logger_instance.info("Starting MSRP file generation.")
        file_url, header_row_from_db, user_email = get_file_location_and_header(file_id_int, logger_instance)
        file_name = os.path.basename(urllib.parse.unquote(file_url))
        local_filename = os.path.join(temp_excel_dir, file_name)

        res = requests.get(file_url, timeout=60)
        res.raise_for_status()
        with open(local_filename, "wb") as f:
            f.write(res.content)

        inferred_header_row = find_header_row_index(local_filename, logger_instance)
        if header_row_from_db is not None:
            try:
                db_header = int(header_row_from_db)
            except (TypeError, ValueError):
                logger_instance.warning(f"Invalid header_row_from_db '{header_row_from_db}'. Falling back to inferred value.")
                db_header = None

            if db_header is not None:
                candidates = {max(0, db_header - 1), max(0, db_header)}
                if inferred_header_row is not None:
                    header_row_value = min(candidates, key=lambda c: abs(c - inferred_header_row))
                else:
                    header_row_value = min(candidates)
            else:
                header_row_value = inferred_header_row if inferred_header_row is not None else 0
        else:
            header_row_value = inferred_header_row if inferred_header_row is not None else 0

        logger_instance.info(f"Using header_row_value={header_row_value} for MSRP template.")
        write_excel_msrp(local_filename, temp_images_dir, grouped_data, header_row_value, target_column, row_offset, logger_instance, populate_images=populate_images, populate_msrp=populate_msrp)

        processed_file_name = f"{Path(file_name).stem}_msrp_{file_type_id}_{timestamp}.xlsx"
        public_url = await upload_file_to_space(local_filename, save_as=f"processed_files/{processed_file_name}", file_id=file_id_int, is_public=True)
        update_file_location_complete(file_id_int, public_url, logger_instance)
        await send_email(
            to_emails=user_email or 'nik@iconluxurygroup.com',
            subject=f'MSRP File Processed: {file_name}',
            file_path=local_filename,
            job_id=file_id
        )
        
        logger_instance.info(f"Successfully completed MSRP job for FileID {file_id}.")

    except Exception as e:
        logger_instance.error(f"FATAL ERROR for MSRP FileID {file_id}: {e}", exc_info=True)
    finally:
        if temp_images_dir and temp_excel_dir:
            await cleanup_temp_dirs([temp_images_dir, temp_excel_dir])

# --- FastAPI Endpoints ---
@app.post("/generate-download-file/")
async def process_file(background_tasks: BackgroundTasks, file_id: int, row_offset: Optional[int] = 0):
    logger.info(f"Received request for FileID: {file_id} with row_offset={row_offset}")
    background_tasks.add_task(generate_download_file, str(file_id), row_offset)
    return {"message": "Processing started. You will be notified upon completion."}

@app.post("/generate-msrp-excel/")
async def process_msrp_file(background_tasks: BackgroundTasks, file_id: int, target_column: str, row_offset: Optional[int] = 0):
    logger.info(f"Received request for MSRP FileID: {file_id} with target_column={target_column} and row_offset={row_offset}")
    background_tasks.add_task(generate_msrp_excel, str(file_id), target_column, row_offset)
    return {"message": "MSRP Processing started. You will be notified upon completion."}

@app.get("/")
def read_root():
    return {"message": "Image Scraper Service is running."}

if __name__ == "__main__":
    import uvicorn
    logger.info("Starting Uvicorn server on http://0.0.0.0:8080")
    uvicorn.run(app, host="0.0.0.0", port=8081)
