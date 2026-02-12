#!/usr/bin/env python3
"""
Script to repair Excel files with corrupted external formula references.
Removes the external links that cause the corruption.
"""

import zipfile
import os
import shutil
from pathlib import Path
import xml.etree.ElementTree as ET


def replace_external_formulas_with_text(sheet_file):
    """
    Replace cells containing external formulas with the text '#REF'.

    Args:
        sheet_file: Path to the worksheet XML file
    """
    try:
        # Parse the XML with namespace handling
        ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

        tree = ET.parse(sheet_file)
        root = tree.getroot()

        # Define namespace
        ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        modified = False

        # Find all cell elements
        for cell in root.findall('.//ss:c', ns):
            # Check if cell has a formula
            formula = cell.find('ss:f', ns)

            if formula is not None:
                formula_text = formula.text or ''

                # Check if formula references external workbook (contains '[' and ']')
                if '[' in formula_text and ']' in formula_text:
                    # Remove the formula element
                    cell.remove(formula)

                    # Set cell type to inline string (str)
                    cell.set('t', 'inlineStr')

                    # Remove any existing value element
                    for v in cell.findall('ss:v', ns):
                        cell.remove(v)

                    # Add inline string value
                    is_elem = ET.SubElement(cell, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
                    t_elem = ET.SubElement(is_elem, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    t_elem.text = '#REF'

                    modified = True

        if modified:
            # Write back to file
            tree.write(sheet_file, encoding='utf-8', xml_declaration=True)
            print(f"  ✓ Updated {sheet_file.name}")

    except Exception as e:
        print(f"  ⚠ Warning: Could not process {sheet_file.name}: {e}")


def repair_excel_file(input_file, output_file=None):
    """
    Repair an Excel file by removing corrupted external links.

    Args:
        input_file: Path to the corrupted Excel file
        output_file: Path for the repaired file (optional, defaults to input_file with _repaired suffix)
    """
    input_path = Path(input_file)

    if output_file is None:
        output_file = input_path.parent / f"{input_path.stem}_repaired{input_path.suffix}"

    output_path = Path(output_file)

    # Create a temporary directory for extraction
    temp_dir = input_path.parent / "temp_excel_repair"
    temp_dir.mkdir(exist_ok=True)

    try:
        print(f"Extracting {input_file}...")
        # Extract the Excel file (which is a ZIP archive)
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Remove external links directory if it exists
        external_links_dir = temp_dir / "xl" / "externalLinks"
        if external_links_dir.exists():
            print(f"Removing external links from {external_links_dir}...")
            shutil.rmtree(external_links_dir)

        # Remove external link references from workbook.xml.rels
        rels_file = temp_dir / "xl" / "_rels" / "workbook.xml.rels"
        if rels_file.exists():
            print(f"Cleaning relationships in {rels_file}...")
            with open(rels_file, 'r', encoding='utf-8') as f:
                content = f.read()

            # Remove external link relationship entries
            import re
            # Remove lines containing externalLinks
            lines = content.split('\n')
            cleaned_lines = [line for line in lines if 'externalLinks' not in line]
            content = '\n'.join(cleaned_lines)

            with open(rels_file, 'w', encoding='utf-8') as f:
                f.write(content)

        # Remove external link references from [Content_Types].xml
        content_types_file = temp_dir / "[Content_Types].xml"
        if content_types_file.exists():
            print(f"Cleaning content types in {content_types_file}...")
            with open(content_types_file, 'r', encoding='utf-8') as f:
                content = f.read()

            # Remove external link override entries
            lines = content.split('\n')
            cleaned_lines = [line for line in lines if 'externalLinks' not in line and 'externalLink' not in line]
            content = '\n'.join(cleaned_lines)

            with open(content_types_file, 'w', encoding='utf-8') as f:
                f.write(content)

        # Replace external formulas with #REF text in all worksheets
        worksheets_dir = temp_dir / "xl" / "worksheets"
        if worksheets_dir.exists():
            print(f"Replacing external formulas with #REF text in worksheets...")
            for sheet_file in worksheets_dir.glob("*.xml"):
                replace_external_formulas_with_text(sheet_file)

        # Create the repaired Excel file
        print(f"Creating repaired file {output_file}...")
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zip_ref.write(file_path, arcname)

        print(f"✓ Successfully repaired Excel file: {output_file}")
        return True

    except Exception as e:
        print(f"✗ Error repairing file: {e}")
        return False

    finally:
        # Clean up temporary directory
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python repair_excel.py <input_file> [output_file]")
        print("\nExample:")
        print("  python repair_excel.py corrupt.xlsx")
        print("  python repair_excel.py corrupt.xlsx repaired.xlsx")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found")
        sys.exit(1)

    success = repair_excel_file(input_file, output_file)
    sys.exit(0 if success else 1)
