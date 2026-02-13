#!/usr/bin/env python3
"""
Test script for Excel repair functionality
"""

import sys
from pathlib import Path
from excel_repair_utils import repair_excel_file, safe_load_workbook
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def test_repair():
    """Test repairing the corrupt.xlsx file"""
    test_file = Path('corrupt.xlsx')

    if not test_file.exists():
        logger.error(f"Test file {test_file} not found")
        return False

    logger.info(f"Testing repair of {test_file}")

    try:
        # Test the repair function
        repaired = repair_excel_file(str(test_file), logger=logger)
        logger.info(f"✓ Repair function completed. Output: {repaired}")

        # Test safe_load_workbook
        logger.info("Testing safe_load_workbook...")
        wb = safe_load_workbook(str(test_file), logger=logger, data_only=True)
        logger.info(f"✓ Loaded workbook with {len(wb.sheetnames)} sheets")

        # Print some info
        ws = wb.active
        logger.info(f"  Active sheet: {ws.title}")
        logger.info(f"  Dimensions: {ws.max_row} rows x {ws.max_column} columns")

        # Check for #REF values
        ref_count = 0
        for row in ws.iter_rows(max_row=10, max_col=10):
            for cell in row:
                if cell.value == '#REF':
                    ref_count += 1

        logger.info(f"  Found {ref_count} #REF text values in first 10x10 cells")

        wb.close()

        logger.info("✓ All tests passed!")
        return True

    except Exception as e:
        logger.error(f"✗ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = test_repair()
    sys.exit(0 if success else 1)
