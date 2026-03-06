#!/bin/bash
set -euo pipefail

# ============================================================
# Super Scraper - Clean Reinstall Helper
#
# Usage:
#   1. Pull the latest source as your normal user
#   2. sudo ./reinstall.sh
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [ "$EUID" -ne 0 ]; then
    echo "ERROR: Please run as root (sudo ./reinstall.sh)"
    exit 1
fi

echo "==> Running clean uninstall..."
"${SCRIPT_DIR}/uninstall.sh"

echo "==> Running fresh install..."
"${SCRIPT_DIR}/install.sh"