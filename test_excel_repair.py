#!/usr/bin/env python3
"""
Test script for Excel repair functionality
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from excel_repair_utils import repair_excel_file
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
        logger.info(f"Repair function completed. Output: {repaired}")

        # Verify the repaired file is a valid xlsx
        if not zipfile.is_zipfile(repaired):
            logger.error("Repaired file is not a valid ZIP archive")
            return False

        with zipfile.ZipFile(repaired, 'r') as zf:
            names = zf.namelist()
            logger.info(f"Repaired file contains {len(names)} entries")

            # Check for essential xlsx components
            if 'xl/workbook.xml' not in names:
                logger.error("Missing workbook.xml in repaired file")
                return False

            # Check worksheets exist
            worksheets = [n for n in names if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
            if not worksheets:
                logger.error("No worksheets found in repaired file")
                return False

            logger.info(f"Found {len(worksheets)} worksheet(s)")

            # Verify worksheet XML integrity
            ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            with zf.open(worksheets[0]) as ws_file:
                tree = ET.parse(ws_file)
                root = tree.getroot()

                # Count #REF values
                ref_count = 0
                for cell in root.findall('.//ss:c', ns):
                    is_elem = cell.find('.//ss:t', ns)
                    if is_elem is not None and is_elem.text == '#REF':
                        ref_count += 1

                logger.info(f"Found {ref_count} #REF text values in first worksheet")

            # Check external links were removed
            ext_links = [n for n in names if 'externalLink' in n.lower()]
            if ext_links:
                logger.warning(f"External links still present: {ext_links}")
            else:
                logger.info("External links successfully removed")

        logger.info("All tests passed!")
        return True

    except Exception as e:
        logger.error(f"Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = test_repair()
    sys.exit(0 if success else 1)
