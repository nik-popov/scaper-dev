import ExcelJS from 'exceljs';
import fs from 'fs/promises';
import { createReadStream, existsSync, statSync } from 'fs';
import readline from 'readline';
import sharp from 'sharp';
import path from 'path';

// Constants
const EMU_PER_PIXEL = 9525; // At 96 DPI
const DPI = 96;
const POINTS_PER_INCH = 72;
const MAX_OFFSET_EMU = 9525000; // ~1000 pixels
const MIN_IMAGE_SIZE = 100; // bytes

/**
 * Convert pixels to EMU (English Metric Units)
 * 1 pixel = 9525 EMU at 96 DPI
 */
function pixelsToEMU(pixels) {
  return Math.round(pixels * EMU_PER_PIXEL);
}

/**
 * Convert points to pixels
 * Points are used for font sizes and row heights
 */
function pointsToPixels(points) {
  return points * DPI / POINTS_PER_INCH;
}

/**
 * Get last non-empty row in a column
 */
function getLastNonEmptyRow(worksheet, column, headerRow) {
  let lastRow = headerRow + 1;
  const col = worksheet.getColumn(column);

  col.eachCell({ includeEmpty: false }, (cell, rowNumber) => {
    if (rowNumber > headerRow + 1 && cell.value) {
      lastRow = Math.max(lastRow, rowNumber);
    }
  });

  return lastRow;
}

/**
 * Excel Bridge class handling all Excel operations via JSON-RPC
 */
class ExcelBridge {
  /**
   * Process a JSON-RPC request
   */
  async processRequest(request) {
    const { method, params, id } = request;

    try {
      let result;
      switch (method) {
        case 'ping':
          result = { status: 'ok', timestamp: new Date().toISOString() };
          break;
        case 'writeExcelDistro':
          result = await this.writeExcelDistro(params);
          break;
        case 'writeExcelMSRP':
          result = await this.writeExcelMSRP(params);
          break;
        case 'writeExcelGeneric':
          result = await this.writeExcelGeneric(params);
          break;
        default:
          throw new Error(`Unknown method: ${method}`);
      }

      return { jsonrpc: '2.0', result, id };
    } catch (error) {
      console.error(`Error in method ${method}:`, error);
      return {
        jsonrpc: '2.0',
        error: {
          code: -32000,
          message: error.message,
          data: error.stack
        },
        id
      };
    }
  }

  /**
   * Write Excel file using distribution template format
   */
  async writeExcelDistro({ templatePath, tempDir, imageData, headerRow, rowOffset = 0 }) {
    console.error(`[writeExcelDistro] Starting - template: ${templatePath}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.worksheets[0];

    console.error(`[writeExcelDistro] Loaded workbook, active worksheet: ${worksheet.name}`);

    // Clear existing images
    const existingImages = worksheet.getImages();
    if (existingImages && existingImages.length > 0) {
      console.error(`[writeExcelDistro] Clearing ${existingImages.length} existing images`);
      // ExcelJS doesn't have removeImages(), we need to manually clear the image list
      worksheet.model.drawing = { rId: worksheet.model.drawing?.rId, drawings: [] };
    }

    // Build image map from temp directory
    const imageMap = {};
    try {
      const files = await fs.readdir(tempDir);
      for (const file of files) {
        const stem = path.parse(file).name;
        if (/^\d+$/.test(stem)) {
          imageMap[parseInt(stem)] = path.join(tempDir, file);
        }
      }
      console.error(`[writeExcelDistro] Found ${Object.keys(imageMap).length} images in ${tempDir}`);
    } catch (error) {
      console.error(`[writeExcelDistro] Warning: Could not read temp dir ${tempDir}: ${error.message}`);
    }

    // Get default row height from template
    const templateRow = worksheet.getRow(headerRow + 2);
    const defaultRowHeight = templateRow.height || 12.75;
    console.error(`[writeExcelDistro] Using template row height: ${defaultRowHeight} points`);

    // Build row mapping
    const rowDataMap = {};
    imageData.forEach(item => {
      rowDataMap[item.ExcelRowID] = item;
    });

    const dataRowIds = Object.keys(rowDataMap).map(id => parseInt(id)).sort((a, b) => a - b);
    const lastNonEmptyRow = getLastNonEmptyRow(worksheet, 'B', headerRow);

    const baseRow = headerRow + rowOffset + 2;
    if (baseRow < 1) {
      throw new Error('Invalid row range: base_row is less than 1');
    }

    // Create row mapping
    const rowIdToRowNum = {};
    dataRowIds.forEach((rowId, idx) => {
      rowIdToRowNum[rowId] = baseRow + idx;
    });

    const maxNeededRow = dataRowIds.length > 0
      ? Math.max(baseRow + dataRowIds.length - 1, lastNonEmptyRow)
      : baseRow;

    console.error(`[writeExcelDistro] Row mapping: base=${baseRow}, dataRows=${dataRowIds.length}, maxNeeded=${maxNeededRow}`);

    // Ensure rows exist
    const currentRowCount = worksheet.rowCount;
    if (currentRowCount < maxNeededRow) {
      console.error(`[writeExcelDistro] Appending ${maxNeededRow - currentRowCount} rows`);
      for (let i = currentRowCount + 1; i <= maxNeededRow; i++) {
        const row = worksheet.getRow(i);
        row.height = defaultRowHeight;
      }
    }

    // Set row height for image rows
    const IMAGE_ROW_HEIGHT = Math.max(defaultRowHeight, 150);

    // Process each data item
    let processedCount = 0;
    for (const rowId of dataRowIds) {
      const rowNum = rowIdToRowNum[rowId];
      const item = rowDataMap[rowId];

      const row = worksheet.getRow(rowNum);
      row.height = IMAGE_ROW_HEIGHT;

      // Insert image if available
      const imagePath = imageMap[rowId];
      if (imagePath) {
        try {
          await this.insertImageWithEMUAnchor(workbook, worksheet, imagePath, rowNum, 0, {
            paddingPoints: 3
          });
          console.error(`[writeExcelDistro] Added image for row ${rowId} at Excel row ${rowNum}`);
        } catch (error) {
          console.error(`[writeExcelDistro] Failed to insert image for row ${rowId}: ${error.message}`);
        }
      }

      // Write metadata (columns B, D, E, H)
      if (rowNum > headerRow + 1) {
        row.getCell(2).value = item.Brand || '';  // Column B
        row.getCell(4).value = item.Style || '';  // Column D
        row.getCell(5).value = item.Color || '';  // Column E
        row.getCell(8).value = item.Category || ''; // Column H
        console.error(`[writeExcelDistro] Wrote metadata for row ${rowId} at Excel row ${rowNum}`);
      }

      processedCount++;
    }

    // Delete extra rows if needed
    if (worksheet.rowCount > maxNeededRow) {
      console.error(`[writeExcelDistro] Deleting ${worksheet.rowCount - maxNeededRow} extra rows`);
      worksheet.spliceRows(maxNeededRow + 1, worksheet.rowCount - maxNeededRow);
    }

    // Set view to A1
    worksheet.views = [{ state: 'normal', topLeftCell: 'A1' }];

    // Save workbook
    console.error(`[writeExcelDistro] Saving workbook to ${templatePath}`);
    await workbook.xlsx.writeFile(templatePath);

    const fileStats = statSync(templatePath);
    console.error(`[writeExcelDistro] File saved successfully, size: ${fileStats.size} bytes`);

    return { success: true, rowsProcessed: processedCount };
  }

  /**
   * Write Excel file using MSRP template format
   */
  async writeExcelMSRP({ templatePath, tempDir, imageData, headerRow, targetColumn, rowOffset, populateImages = true, populateMSRP = true }) {
    console.error(`[writeExcelMSRP] Starting - template: ${templatePath}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.worksheets[0];

    console.error(`[writeExcelMSRP] Loaded workbook, active worksheet: ${worksheet.name}`);

    // Clear existing images
    if (populateImages) {
      const existingImages = worksheet.getImages();
      if (existingImages && existingImages.length > 0) {
        console.error(`[writeExcelMSRP] Clearing ${existingImages.length} existing images`);
        worksheet.model.drawing = { rId: worksheet.model.drawing?.rId, drawings: [] };
      }
    }

    // Build image map
    const imageMap = {};
    if (populateImages) {
      try {
        const files = await fs.readdir(tempDir);
        for (const file of files) {
          const stem = path.parse(file).name;
          if (/^\d+$/.test(stem)) {
            imageMap[parseInt(stem)] = file;
          }
        }
        console.error(`[writeExcelMSRP] Found ${Object.keys(imageMap).length} images in ${tempDir}`);
      } catch (error) {
        console.error(`[writeExcelMSRP] Warning: Could not read temp dir ${tempDir}: ${error.message}`);
      }
    }

    // Validate target column
    if (populateMSRP && !/^[A-Z]+$/.test(targetColumn)) {
      throw new Error(`Invalid target_column: ${targetColumn}. Must be a valid Excel column letter (e.g., 'A', 'B', 'AA').`);
    }

    // Get default row height
    const templateRow = worksheet.getRow(headerRow + 2);
    const defaultRowHeight = templateRow.height || 12.75;
    console.error(`[writeExcelMSRP] Using template row height: ${defaultRowHeight} points`);

    // Build row mapping
    const rowDataMap = {};
    imageData.forEach(item => {
      rowDataMap[item.ExcelRowID] = item;
    });

    const dataRowIds = Object.keys(rowDataMap).map(id => parseInt(id)).sort((a, b) => a - b);
    const lastNonEmptyRow = getLastNonEmptyRow(worksheet, 'B', headerRow);

    const baseRow = headerRow + rowOffset + 2;
    if (baseRow < 1) {
      throw new Error('Invalid row range: base_row is less than 1');
    }

    // Create row mapping
    const rowIdToRowNum = {};
    dataRowIds.forEach((rowId, idx) => {
      rowIdToRowNum[rowId] = baseRow + idx;
    });

    const maxNeededRow = dataRowIds.length > 0
      ? Math.max(baseRow + dataRowIds.length - 1, lastNonEmptyRow)
      : baseRow;

    console.error(`[writeExcelMSRP] Row mapping: base=${baseRow}, dataRows=${dataRowIds.length}, maxNeeded=${maxNeededRow}`);

    // Ensure rows exist
    const currentRowCount = worksheet.rowCount;
    if (currentRowCount < maxNeededRow) {
      console.error(`[writeExcelMSRP] Appending ${maxNeededRow - currentRowCount} rows`);
      for (let i = currentRowCount + 1; i <= maxNeededRow; i++) {
        const row = worksheet.getRow(i);
        row.height = defaultRowHeight;
      }
    }

    // Process each data item
    let processedCount = 0;
    for (const rowId of dataRowIds) {
      const rowNum = rowIdToRowNum[rowId];
      const item = rowDataMap[rowId];

      const row = worksheet.getRow(rowNum);
      row.height = defaultRowHeight;

      // Insert image if enabled
      if (populateImages && rowId in imageMap) {
        const imagePath = path.join(tempDir, imageMap[rowId]);
        try {
          // Insert image and get its dimensions
          const { imgHeightPoints } = await this.insertImageWithEMUAnchor(workbook, worksheet, imagePath, rowNum, 0, {
            paddingPoints: 2
          });

          // Adjust row height if image is taller
          if (imgHeightPoints && imgHeightPoints > defaultRowHeight) {
            row.height = Math.max(defaultRowHeight, imgHeightPoints);
          }

          console.error(`[writeExcelMSRP] Added image for row ${rowId} at Excel row ${rowNum}, height=${row.height}`);
        } catch (error) {
          console.error(`[writeExcelMSRP] Failed to insert image for row ${rowId}: ${error.message}`);
        }
      }

      // Write MSRP value if enabled
      if (populateMSRP && item) {
        const msrpValue = item.MSRP || '';
        row.getCell(targetColumn).value = msrpValue;
        console.error(`[writeExcelMSRP] Wrote MSRP '${msrpValue}' for row ${rowId} at ${targetColumn}${rowNum}`);
      }

      processedCount++;
    }

    // Delete extra rows if needed
    if (worksheet.rowCount > maxNeededRow) {
      console.error(`[writeExcelMSRP] Deleting ${worksheet.rowCount - maxNeededRow} extra rows`);
      worksheet.spliceRows(maxNeededRow + 1, worksheet.rowCount - maxNeededRow);
    }

    // Set view to A1
    worksheet.views = [{ state: 'normal', topLeftCell: 'A1' }];

    // Save workbook
    console.error(`[writeExcelMSRP] Saving workbook to ${templatePath}`);
    await workbook.xlsx.writeFile(templatePath);

    const fileStats = statSync(templatePath);
    console.error(`[writeExcelMSRP] File saved successfully, size: ${fileStats.size} bytes`);

    return { success: true, rowsProcessed: processedCount };
  }

  /**
   * Write Excel file using generic template format
   */
  async writeExcelGeneric({ templatePath, tempDir, imageData, headerRow, rowOffset, fileTypeId = null }) {
    console.error(`[writeExcelGeneric] Starting - template: ${templatePath}, fileTypeId: ${fileTypeId}`);

    // For now, generic follows the same pattern as distro
    // Can be customized based on fileTypeId if needed
    return await this.writeExcelDistro({ templatePath, tempDir, imageData, headerRow, rowOffset });
  }

  /**
   * Insert image with precise EMU-based anchoring
   * Matches openpyxl's OneCellAnchor behavior
   */
  async insertImageWithEMUAnchor(workbook, worksheet, imagePath, rowNum, colNum, options) {
    // Validate image file exists
    if (!existsSync(imagePath)) {
      throw new Error(`Image file does not exist: ${imagePath}`);
    }

    const fileStats = statSync(imagePath);
    if (fileStats.size < MIN_IMAGE_SIZE) {
      throw new Error(`Image file too small (${fileStats.size} bytes): ${imagePath}`);
    }

    // Get image dimensions using Sharp
    let metadata;
    try {
      metadata = await sharp(imagePath).metadata();
    } catch (error) {
      throw new Error(`Failed to read image metadata: ${error.message}`);
    }

    const imgWidthPixels = metadata.width;
    const imgHeightPixels = metadata.height;

    if (!imgWidthPixels || !imgHeightPixels || imgWidthPixels <= 0 || imgHeightPixels <= 0) {
      throw new Error(`Invalid image dimensions: ${imgWidthPixels}x${imgHeightPixels}`);
    }

    // Read actual cell dimensions from the worksheet
    const col = worksheet.getColumn(colNum + 1); // ExcelJS columns are 1-based
    const row = worksheet.getRow(rowNum);
    const colWidthChars = col.width || 8.43;        // Excel character-width units
    const rowHeightPoints = row.height || 15;        // Points

    // Excel character-width → pixels:  chars * 7 + 5
    const cellWidthPixels = Math.round(colWidthChars * 7 + 5);
    // Points → pixels
    const cellHeightPixels = pointsToPixels(rowHeightPoints);

    // Apply padding
    const paddingPx = options.paddingPoints ? pointsToPixels(options.paddingPoints) : 4;
    const availWidth = Math.max(1, cellWidthPixels - paddingPx * 2);
    const availHeight = Math.max(1, cellHeightPixels - paddingPx * 2);

    // Scale image to fit within the available cell area (preserve aspect ratio)
    const scaleX = availWidth / imgWidthPixels;
    const scaleY = availHeight / imgHeightPixels;
    const scale = Math.min(scaleX, scaleY, 1);  // never upscale

    const finalWidth = Math.round(imgWidthPixels * scale);
    const finalHeight = Math.round(imgHeightPixels * scale);
    const imgHeightPoints = finalHeight * POINTS_PER_INCH / DPI;

    // Calculate centering offsets in pixels
    const xOffsetPixels = Math.max(0, (cellWidthPixels - finalWidth) / 2);
    const yOffsetPixels = Math.max(0, (cellHeightPixels - finalHeight) / 2);

    // Convert offsets to EMU for nativeColOff / nativeRowOff
    const xOffsetEMU = Math.round(xOffsetPixels * EMU_PER_PIXEL);
    const yOffsetEMU = Math.round(yOffsetPixels * EMU_PER_PIXEL);

    console.error(
      `[insertImage] Row ${rowNum}: img=${imgWidthPixels}x${imgHeightPixels}px, ` +
      `scaled=${finalWidth}x${finalHeight}px (${(scale * 100).toFixed(0)}%), ` +
      `cell=${Math.round(cellWidthPixels)}x${Math.round(cellHeightPixels)}px, ` +
      `offsets=(${xOffsetPixels.toFixed(1)}, ${yOffsetPixels.toFixed(1)})px`
    );

    // Read image file
    const imageBuffer = await fs.readFile(imagePath);
    const extension = path.extname(imagePath).substring(1).toLowerCase();

    // Add image to workbook
    const imageId = workbook.addImage({
      buffer: imageBuffer,
      extension: extension === 'jpg' ? 'jpeg' : extension
    });

    // ExcelJS ext expects PIXELS (it multiplies by 9525 internally)
    // ExcelJS tl with col/row ignores colOff/rowOff — use nativeCol/nativeColOff for EMU offsets
    worksheet.addImage(imageId, {
      tl: {
        nativeCol: colNum,
        nativeColOff: xOffsetEMU,
        nativeRow: rowNum - 1,
        nativeRowOff: yOffsetEMU
      },
      ext: {
        width: finalWidth,
        height: finalHeight
      },
      editAs: 'oneCell'
    });

    return { imgHeightPoints };
  }
}

/**
 * Main process loop - JSON-RPC over stdin/stdout
 */
async function main() {
  const bridge = new ExcelBridge();

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    terminal: false
  });

  console.error('[ExcelJS Bridge] Started and ready for requests');

  rl.on('line', async (line) => {
    try {
      const request = JSON.parse(line);
      const response = await bridge.processRequest(request);
      console.log(JSON.stringify(response));
    } catch (error) {
      const errorResponse = {
        jsonrpc: '2.0',
        error: {
          code: -32700,
          message: 'Parse error',
          data: error.message
        },
        id: null
      };
      console.log(JSON.stringify(errorResponse));
    }
  });

  rl.on('close', () => {
    console.error('[ExcelJS Bridge] Stdin closed, exiting');
    process.exit(0);
  });
}

main().catch(error => {
  console.error('[ExcelJS Bridge] Fatal error:', error);
  process.exit(1);
});
