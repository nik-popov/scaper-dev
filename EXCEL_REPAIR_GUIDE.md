# Excel Repair System - Implementation Guide

## Overview

Your system now automatically detects and repairs corrupted Excel files with external link and linked data type issues. This resolves the error:

```
Excel completed file level validation and repair. Some parts of this workbook may have been repaired or discarded.
Repaired Records: External formula reference from /xl/externalLinks/externalLink1.xml part
```

## What Was Implemented

### 1. **Excel Repair Utility Module** (`excel_repair_utils.py`)

Core utilities for repairing Excel files:

- `repair_excel_file()` - Repairs corrupted Excel files by:
  - Removing broken external link XML files
  - Cleaning relationship references
  - Replacing external formulas with "#REF" text
  - Fixing linked data type columns (FIELDVALUE, FILTERXML, etc.)

- `safe_load_workbook()` - Drop-in replacement for `openpyxl.load_workbook()` that:
  - Automatically detects corruption when loading fails
  - Repairs the file if needed
  - Returns the workbook ready to use

### 2. **FastAPI Endpoint** (`/reformatExcel`)

New REST endpoint in [main.py](main.py:1875) for on-demand Excel repair:

```python
POST /reformatExcel
Body: { "file_url": "https://..." }

Response: {
  "success": true,
  "message": "Excel file repaired successfully",
  "s3_url": "https://...",
  "r2_url": "https://...",
  "file_id": "uuid"
}
```

### 3. **Automatic Repair Integration**

Updated [main.py:698](main.py:698) to use `safe_load_workbook()` instead of `load_workbook()`:

```python
# OLD: wb = load_workbook(cleaned_filename)
# NEW: wb = safe_load_workbook(cleaned_filename, logger=logger_instance)
```

This ensures corrupted files are automatically repaired during processing.

### 4. **Command-Line Tools**

#### Standalone Repair Script
```bash
python repair_excel.py corrupt.xlsx [output.xlsx]
```

#### Wrapper Script
```bash
./reformat_excel.sh corrupt.xlsx
```

#### Test Suite
```bash
python test_excel_repair.py
```

### 5. **Claude Code Skill** (`/reformat`)

Created skill at `~/.claude/skills/reformat.md` for easy invocation:

```
/reformat=excel <filename>
```

## How It Works

### Detection & Repair Process

1. **Extract**: Unzip the Excel file (it's a ZIP archive)
2. **Detect**: Check for corruption indicators:
   - External links directory exists
   - External link references in relationships
   - Formulas with `[ExternalWorkbook]` syntax
   - Linked data type functions (FIELDVALUE, FILTERXML, etc.)
3. **Repair**:
   - Remove external link XML files
   - Clean metadata references
   - Replace problematic formulas with "#REF" text
4. **Rebuild**: Create clean Excel file

### Test Results

On your `corrupt.xlsx` file:
- ✓ Detected and removed external links directory
- ✓ Cleaned relationship references
- ✓ Replaced **331 external formula references** with "#REF" text
- ✓ File now opens without corruption warnings

## Usage Examples

### In Your Python Code

```python
from excel_repair_utils import safe_load_workbook

# Automatic repair if needed
wb = safe_load_workbook('myfile.xlsx', logger=logger)
ws = wb.active
# ... process normally
```

### As API Endpoint

```bash
curl -X POST https://your-domain.com/reformatExcel \
  -H "Content-Type: application/json" \
  -d '{"file_url": "https://example.com/corrupt.xlsx"}'
```

### Command Line

```bash
# Repair a file
python repair_excel.py input.xlsx output.xlsx

# Test repair functionality
python test_excel_repair.py
```

## Integration Points

Your existing backend automatically benefits from this system:

1. **File Upload Processing**: When users upload Excel files, corrupted files are detected and repaired automatically
2. **Image Extraction**: The system in [main.py](main.py) now handles corrupted files gracefully
3. **Database Import**: Files are cleaned before data is inserted into `utb_ExcelParser`

## What Gets Fixed

### External Formula References
```excel
Before: =[OtherWorkbook.xlsx]Sheet1!A1
After:  #REF (as text)
```

### Linked Data Types
```excel
Before: =FIELDVALUE(A1, "Price")
After:  #REF (as text)
```

### Error Cells
```excel
Before: #REF! (formula error)
After:  #REF (plain text)
```

## Files Modified

- ✓ [main.py](main.py) - Added `/reformatExcel` endpoint and integrated `safe_load_workbook`
- ✓ [excel_repair_utils.py](excel_repair_utils.py) - New repair utilities module
- ✓ [repair_excel.py](repair_excel.py) - Enhanced standalone repair script
- ✓ [test_excel_repair.py](test_excel_repair.py) - Test suite
- ✓ [~/.claude/skills/reformat.md](~/.claude/skills/reformat.md) - Claude Code skill

## Next Steps

1. **Deploy**: Restart your FastAPI service to load the new endpoint
2. **Test**: Try uploading a corrupted Excel file
3. **Monitor**: Check logs for automatic repair messages
4. **Optional**: Update frontend to call `/reformatExcel` before processing

## Support

If you encounter issues:
1. Check logs for repair details
2. Run `python test_excel_repair.py` to verify setup
3. Use `python repair_excel.py <file>` to manually test specific files
