#!/bin/bash
# Wrapper script for Excel reformatting skill

if [ $# -eq 0 ]; then
    echo "Usage: reformat_excel.sh <excel_file>"
    exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
python "${SCRIPT_DIR}/repair_excel.py" "$@"
