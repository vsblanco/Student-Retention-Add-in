// [2025-12-03] Version 2.2 - Advanced debounce with stale-request cancellation to prevent backlog
import { getCanonicalColIdx, canonicalHeaderMap } from './CanonicalMap.jsx';

export function normalize(val) {
    if (val === undefined || val === null) return "";
    // Convert to string, remove all spaces globally, and convert to lowercase.
    return String(val).replace(/\s/g, '').toLowerCase();
}

// Helper: extract Hyperlink from Excel HYPERLINK formula or plain URL
function extractHyperlink(formulaOrValue) {
  if (!formulaOrValue) return null;
  if (typeof formulaOrValue !== 'string') return null;
  const value = formulaOrValue.trim();
  const hyperlinkRegex = /=\s*HYPERLINK\s*\(\s*["']?([^"',)]+)["']?\s*,/i;
  const match = value.match(hyperlinkRegex);
  if (match && match[1]) {
    return match[1].trim();
  }
  if (/^https?:\/\//i.test(value)) {
    return value;
  }
  return null;
}

// Helper: convert 0-based column index to A1 column letters (0 -> A, 25 -> Z, 26 -> AA)
function colIndexToLetter(colIndex) {
    if (typeof colIndex !== 'number' || colIndex < 0) return '';
    let letters = '';
    let n = colIndex + 1;
    while (n > 0) {
        const rem = (n - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}

/**
 * Edits a row in an Excel worksheet by unique ID.
 * ... (unchanged)
 */
export async function editRow(sheet, columnId, rowId, newData) {
    console.log("ExcelAPI.editRow: start", { sheet, columnId, rowId });
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        console.log("ExcelAPI.editRow: finished - invalid parameters");
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }

    try {
        let changeVerified = false;
        let errorMsg = "";
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet);
            const usedRange = worksheet.getUsedRange();
            usedRange.load(["values", "columnCount", "rowCount"]);
            await context.sync();

            const headers = usedRange.values[0];

            // Use canonical matching to find ID column (handles aliases and normalization)
            const idColIdx = getCanonicalColIdx(headers, columnId, canonicalHeaderMap);

            if (idColIdx === -1) {
                throw new Error(`ID column "${columnId}" not found.`);
            }

            // Find row index by ID
            let rowIdx = -1;
            for (let i = 1; i < usedRange.rowCount; i++) {
                if (usedRange.values[i][idColIdx] == rowId) {
                    rowIdx = i;
                    break;
                }
            }
            if (rowIdx === -1) {
                throw new Error(`Row with ID "${rowId}" not found.`);
            }

            // Update specified columns
            for (const [colName, value] of Object.entries(newData)) {
                 // Resolve column index via canonical mapping (handles aliases)
                 const colIdx = getCanonicalColIdx(headers, colName, canonicalHeaderMap);
                 if (colIdx === -1) continue; // skip unknown columns
                 worksheet.getCell(rowIdx, colIdx).values = [[value]];
             }
             await context.sync();

             // Double check
             const verifyRange = worksheet.getRangeByIndexes(rowIdx, 0, 1, usedRange.columnCount);
             verifyRange.load("values");
             await context.sync();
             const verifyValues = verifyRange.values[0];
             changeVerified = Object.entries(newData).every(([colName, value]) => {
                 const colIdx = getCanonicalColIdx(headers, colName, canonicalHeaderMap);
                 return colIdx !== -1 && verifyValues[colIdx] == value;
             });
             if (!changeVerified) {
                 errorMsg = "Double check failed: Row was not updated as expected.";
             }
         });
         if (!changeVerified) {
             return { success: false, message: errorMsg };
         }
        console.log("ExcelAPI.editRow: finished - success");
        return { success: true, message: "Row updated successfully." };
    } catch (error) {
        console.log("ExcelAPI.editRow: finished - error", error?.message);
        return { success: false, message: error.message };
    }
}

/**
 * Inserts a new row into an Excel worksheet.
 * ... (unchanged)
 */
export async function insertRow(sheet, newData) {
    console.log("ExcelAPI.insertRow: start", { sheet });
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        console.log("ExcelAPI.insertRow: finished - invalid parameters");
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }
    try {
        let changeVerified = false;
        let errorMsg = "";
        await Excel.run(async (context) => {
             const worksheet = context.workbook.worksheets.getItem(sheet);
             const usedRange = worksheet.getUsedRange();
             usedRange.load(["values", "columnCount", "rowCount"]);
             await context.sync();

             const headers = usedRange.values[0];

             const newRow = Array(usedRange.columnCount).fill("");
             for (const [colName, value] of Object.entries(newData)) {
                 // Resolve column by canonical mapping
                 const colIdx = getCanonicalColIdx(headers, colName, canonicalHeaderMap);
                  if (colIdx !== -1) {
                      newRow[colIdx] = value;
                 }
             }
              worksheet.getRangeByIndexes(usedRange.rowCount, 0, 1, usedRange.columnCount).values = [newRow];
              await context.sync();

              // Double check
              const verifyRange = worksheet.getRangeByIndexes(usedRange.rowCount, 0, 1, usedRange.columnCount);
              verifyRange.load("values");
              await context.sync();
              const verifyValues = verifyRange.values[0];
              changeVerified = Object.entries(newData).every(([colName, value]) => {
                  const colIdx = getCanonicalColIdx(headers, colName, canonicalHeaderMap);
                  return colIdx !== -1 && verifyValues[colIdx] == value;
              });
              if (!changeVerified) {
                 errorMsg = "Double check failed: Row was not inserted as expected.";
              }
          });
          if (!changeVerified) {
              return { success: false, message: errorMsg };
          }
        console.log("ExcelAPI.insertRow: finished - success");
        return { success: true, message: "Row inserted successfully." };
    } catch (error) {
        console.log("ExcelAPI.insertRow: finished - error", error?.message);
        return { success: false, message: error.message };
    }
}

/**
 * Deletes a row in an Excel worksheet by unique ID.
 * ... (unchanged)
 */
export async function deleteRow(sheet, columnId, rowId) {
    console.log("ExcelAPI.deleteRow: start", { sheet, columnId, rowId });
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
         return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
     }
     try {
        let changeVerified = false;
        let errorMsg = "";
        await Excel.run(async (context) => {
             const worksheet = context.workbook.worksheets.getItem(sheet);
             const usedRange = worksheet.getUsedRange();
             usedRange.load(["values", "columnCount", "rowCount"]);
             await context.sync();

             const headers = usedRange.values[0];

             const idColIdx = getCanonicalColIdx(headers, columnId, canonicalHeaderMap);
             if (idColIdx === -1) {
                 throw new Error(`ID column "${columnId}" not found.`);
             }

             // Find row index by ID
             let rowIdx = -1;
             for (let i = 1; i < usedRange.rowCount; i++) {
                  if (usedRange.values[i][idColIdx] == rowId) {
                      rowIdx = i;
                      break;
                  }
              }
              if (rowIdx === -1) {
                  throw new Error(`Row with ID "${rowId}" not found.`);
              }

              worksheet.getRangeByIndexes(rowIdx, 0, 1, usedRange.columnCount).delete(Excel.DeleteShiftDirection.up);
              await context.sync();

              // Double check
              const verifyRange = worksheet.getUsedRange();
              verifyRange.load("values");
              await context.sync();
              changeVerified = verifyRange.values.slice(1).every(row => row[idColIdx] != rowId);
              if (!changeVerified) {
                  errorMsg = "Double check failed: Row was not deleted as expected.";
              }
          });
          if (!changeVerified) {
              return { success: false, message: errorMsg };
          }
         console.log("ExcelAPI.deleteRow: finished - success");
         return { success: true, message: "Row deleted successfully." };
     } catch (error) {
         console.log("ExcelAPI.deleteRow: finished - error", error?.message);
         return { success: false, message: error.message };
     }
}

/**
 * Registers a callback for Excel worksheet selection changes.
 * * Uses debounce and a stale-request check to handle rapid clicking without backlog.
 * * @param {function} callback - Function to run when selection changes.
 * @param {object} COLUMN_ALIASES - Optional mapping of canonical column names to alias names.
 * @returns {Promise<{remove: function}>} - Returns an object with a remove() function to unregister the event.
 */
export async function onSelectionChanged(callback, COLUMN_ALIASES = null) {
    console.log("ExcelAPI.onSelectionChanged: start");
    let eventHandler;
    let debounceTimer = null;
    const DEBOUNCE_DELAY = 250; // ms

    // Tracker for the latest event ID to handle "stale" requests
    let latestRequestId = 0;

    await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        
        eventHandler = worksheet.onSelectionChanged.add(async (eventArgs) => {
            // 1. Increment ID immediately on every raw event
            latestRequestId++;
            const currentRequestId = latestRequestId;

            // 2. Clear previous timer
            if (debounceTimer) {
                clearTimeout(debounceTimer);
            }

            // 3. Set new timer
            debounceTimer = setTimeout(async () => {
                // 4. Stale check: If a newer event came in while waiting, abort this one
                if (currentRequestId !== latestRequestId) {
                    return; 
                }

                try {
                    // 5. Run the data fetch in its OWN isolated context to prevent queueing on the registration context
                    await Excel.run(async (newContext) => {
                        const activeSheet = newContext.workbook.worksheets.getActiveWorksheet();
                        // Get whatever is CURRENTLY selected (resolves race conditions better than eventArgs.address for rapid clicks)
                        const currentSelection = newContext.workbook.getSelectedRange();
                        currentSelection.load("address");
                        
                        // Delegate to loadRange
                        const payload = await loadRange(newContext, activeSheet, currentSelection, COLUMN_ALIASES);
                        
                        // 6. FINAL Stale check: If a newer event happened while we were fetching data, discard this result.
                        // This prevents "ghost" updates from old clicks appearing after the new content.
                        if (currentRequestId !== latestRequestId) {
                            console.log("ExcelAPI: Discarding stale selection data for ID:", currentRequestId);
                            return; 
                        }

                        if (!payload || !payload.success) {
                            callback({});
                            return;
                        }

                        // invoke callback with fresh data
                        callback({
                            address: payload.address,
                            data: payload.data || {},
                            values: payload.values || [],
                            rowCount: payload.rowCount || 1,
                            allRows: payload.allRows || [],
                            hiddenRowCount: payload.hiddenRowCount || 0
                        });
                    });
                } catch (err) {
                    console.error("onSelectionChanged processing error:", err);
                }
            }, DEBOUNCE_DELAY);
        });
        await context.sync();
    });

    console.log("ExcelAPI.onSelectionChanged: finished - handler registered");
    
    return {
        remove: async () => {
            if (debounceTimer) {
                clearTimeout(debounceTimer);
            }
            if (eventHandler) {
                await Excel.run(async (context) => {
                    eventHandler.remove();
                    await context.sync();
                });
            }
        }
    };
}

// New: reusable loader for a given range address within an existing Excel.run context.
// Supports passing either a rangeAddress string OR a Range object directly as the third argument.
export async function loadRange(context, worksheet, rangeOrAddress, COLUMN_ALIASES = null) {
    let selRange;
    if (typeof rangeOrAddress === 'string') {
        selRange = worksheet.getRange(rangeOrAddress);
    } else {
        // Assume it's a Range object already linked to 'context'
        selRange = rangeOrAddress;
    }
    
    selRange.load(["address", "rowIndex", "rowCount"]);

    // Load usedRange headers and position
    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load(["values", "formulas", "rowIndex", "columnIndex", "columnCount", "isNullObject"]);
    await context.sync();

    if (!usedRange || usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
        // If empty sheet, return graceful empty
        return { success: false, message: "No used range / headers found.", address: selRange.address };
    }

    // Normalize raw header strings
    const rawHeaders = usedRange.values[0].map(h => (h === undefined || h === null) ? "" : String(h));
    const usedColIndex = (typeof usedRange.columnIndex === "number") ? usedRange.columnIndex : 0;
    const usedColCount = usedRange.columnCount || (rawHeaders.length);

    // Ensure selRange.rowIndex is available
    const selectedRowIndex = selRange.rowIndex;
    const selectedRowCount = selRange.rowCount || 1;

    // Read the full selected range across the used columns (handles multiple rows)
    const rowRange = worksheet.getRangeByIndexes(selectedRowIndex, usedColIndex, selectedRowCount, usedColCount);
    rowRange.load("values");
    await context.sync();

    // Check which rows are hidden to filter them out
    const rowVisibility = [];
    for (let r = 0; r < selectedRowCount; r++) {
        const row = worksheet.getRangeByIndexes(selectedRowIndex + r, usedColIndex, 1, 1);
        row.load("rowHidden");
        rowVisibility.push(row);
    }
    await context.sync();

    const usedRangeStartRow = (typeof usedRange.rowIndex === "number") ? usedRange.rowIndex : 0;

    // Helper function to build row object from values
    const buildRowObject = (rowValues, relativeRowIdx) => {
        const formulasRow = (usedRange.formulas && usedRange.formulas[relativeRowIdx]) ? usedRange.formulas[relativeRowIdx] : [];
        const rowObj = {};

        for (let c = 0; c < usedColCount; c++) {
            const headerRaw = (rawHeaders[c] !== undefined && rawHeaders[c] !== null) ? String(rawHeaders[c]) : "";
            let headerKey = headerRaw;

            if (COLUMN_ALIASES && typeof COLUMN_ALIASES === "object") {
                const normalizedHeader = normalize(headerRaw);
                const match = Object.keys(COLUMN_ALIASES).find((canonical) => {
                    if (!canonical) return false;
                    const normalizedCanonical = normalize(canonical);
                    if (normalizedCanonical === normalizedHeader) return true;
                    const aliasesRaw = COLUMN_ALIASES[canonical];
                    const aliases = Array.isArray(aliasesRaw)
                        ? aliasesRaw
                        : (typeof aliasesRaw === "string" ? [aliasesRaw] : []);
                    return aliases.some(a => normalize(a) === normalizedHeader);
                });
                if (match) {
                    headerKey = match;
                }
            }

            if (!headerKey || headerKey === "") {
                headerKey = `Column${usedColIndex + c}`;
            }

            // Prefer formula text from usedRange.formulas for hyperlink extraction.
            let cellVal = (rowValues[c] !== undefined) ? rowValues[c] : null;
            const formulaCell = (formulasRow && formulasRow[c] !== undefined) ? formulasRow[c] : null;
            if (typeof formulaCell === 'string' && formulaCell.trim().startsWith('=')) {
                const link = extractHyperlink(formulaCell);
                if (link) {
                    cellVal = link;
                }
            } else if (typeof cellVal === 'string' && cellVal.trim() !== '') {
                // Fallback: if value itself looks like a formula, handle it (rare if formulas were loaded)
                const rawTrim = cellVal.trim();
                if (rawTrim.startsWith('=')) {
                    const link = extractHyperlink(rawTrim);
                    if (link) {
                        cellVal = link;
                    }
                }
            }

            rowObj[headerKey] = cellVal;
        }

        return rowObj;
    };

    // Build all row objects, filtering out hidden rows
    const allRows = [];
    for (let r = 0; r < selectedRowCount; r++) {
        // Skip hidden rows
        if (rowVisibility[r] && rowVisibility[r].rowHidden) {
            continue;
        }

        const absoluteRowIndex = selectedRowIndex + r;
        const relativeRowIdx = absoluteRowIndex - usedRangeStartRow;
        const rowValues = (rowRange.values && rowRange.values[r]) ? rowRange.values[r] : [];
        const rowObj = buildRowObject(rowValues, relativeRowIdx);
        allRows.push(rowObj);
    }

    // Build headers index map
    const headerIndexMap = {};
    rawHeaders.forEach((h, i) => { headerIndexMap[h] = i; });

    // For backward compatibility, keep first visible row as 'data' and first visible row values as 'values'
    // Find the first visible row index
    let firstVisibleRowIdx = 0;
    for (let r = 0; r < selectedRowCount; r++) {
        if (!rowVisibility[r] || !rowVisibility[r].rowHidden) {
            firstVisibleRowIdx = r;
            break;
        }
    }

    const firstRowValues = (rowRange.values && rowRange.values[firstVisibleRowIdx]) ? rowRange.values[firstVisibleRowIdx] : [];
    const firstRowData = allRows.length > 0 ? allRows[0] : {};
    const firstRelativeRowIdx = (selectedRowIndex + firstVisibleRowIdx) - usedRangeStartRow;
    const formulasRow = (usedRange.formulas && usedRange.formulas[firstRelativeRowIdx]) ? usedRange.formulas[firstRelativeRowIdx] : [];

    return {
        success: true,
        address: selRange.address,
        rowIndex: selectedRowIndex,
        rowCount: allRows.length,  // Return count of visible rows only
        originalRowCount: selectedRowCount,  // Original selection count (including hidden)
        hiddenRowCount: selectedRowCount - allRows.length,  // Number of hidden rows
        startCol: usedColIndex,
        columnCount: usedColCount,
        headers: headerIndexMap,
        values: firstRowValues,
        formulas: formulasRow,
        data: firstRowData,
        allRows: allRows  // All visible (non-hidden) rows only
    };
}

/**
 * Checks if a row exists in the worksheet by unique ID.
 * ... (unchanged)
 */
export async function checkRow(sheet, columnId, rowId) {
     if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
         throw new Error("Parameter 'sheet' is required and must be a non-empty string.");
     }
     console.log("ExcelAPI.checkRow: start", { sheet, columnId, rowId });
     try {
         let exists = false;
         await Excel.run(async (context) => {
             const worksheet = context.workbook.worksheets.getItem(sheet);
             const usedRange = worksheet.getUsedRange();
             usedRange.load(["values", "rowCount", "columnCount"]);
             await context.sync();

             const headers = usedRange.values[0];
             const idColIdx = getCanonicalColIdx(headers, columnId, canonicalHeaderMap);
             if (idColIdx === -1) {
                 exists = false;
                 return;
             }
             for (let i = 1; i < usedRange.rowCount; i++) {
                 if (usedRange.values[i][idColIdx] == rowId) { 
                    exists = true;
                    break;
                 }
             }
         });
         console.log("ExcelAPI.checkRow: finished - result", exists);
         return exists;
     } catch (error) {
         console.error("ExcelAPI.checkRow: finished - error", error && error.message ? error.message : error);
         throw new Error(error.message || String(error));
     }
}

/**
 * Highlights a row in the active worksheet.
 * ... (unchanged)
 */
export async function highlightRow(rowIndex, startCol, colCount, color = 'yellow') {
  if (typeof window.Excel === "undefined") return;
  if (typeof rowIndex !== 'number' || typeof startCol !== 'number' || typeof colCount !== 'number') return;
  if (colCount <= 0) return;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let leftStartCol = startCol - (colCount - 1);
      if (leftStartCol < 0) leftStartCol = 0;
      const actualColCount = (startCol - leftStartCol) + 1;
      if (actualColCount <= 0) return;

      const highlightRange = sheet.getRangeByIndexes(rowIndex, leftStartCol, 1, actualColCount);
      highlightRange.format.fill.color = color;
      await context.sync();
    });
  } catch (_) { }
}

// NEW: clear fill for a row range (inverse of highlightRow)
export async function clearRowFill(rowIndex, startCol, colCount) {
  if (typeof window.Excel === "undefined") return;
  if (typeof rowIndex !== 'number' || typeof startCol !== 'number' || typeof colCount !== 'number') return;
  if (colCount <= 0) return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let leftStartCol = startCol - (colCount - 1);
      if (leftStartCol < 0) leftStartCol = 0;
      const actualColCount = (startCol - leftStartCol) + 1;
      if (actualColCount <= 0) return;
      const targetRange = sheet.getRangeByIndexes(rowIndex, leftStartCol, 1, actualColCount);
      targetRange.format.fill.clear();
      await context.sync();
    });
  } catch (_) { }
}

/**
 * Loads a worksheet's used range and returns headers, values and dimensions.
 * ... (unchanged)
 */
export async function loadSheet(sheet, identifierColumn = null, identifierRow = null) {
    console.log("ExcelAPI.loadSheet: start", { sheet }, { identifierColumn });
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        console.log("ExcelAPI.loadSheet: finished - invalid parameters");
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }

    try {
        let payload = null;
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItemOrNullObject(sheet);
            worksheet.load("isNullObject");
            await context.sync();

            if (worksheet.isNullObject) {
                throw new Error(`Worksheet "${sheet}" not found.`);
            }

            const usedRange = worksheet.getUsedRangeOrNullObject();
            usedRange.load(["values", "rowCount", "columnCount", "address", "isNullObject"]);
            await context.sync();

            if (usedRange.isNullObject) {
                payload = { rowCount: 0, columnCount: 0, headers: {}, data: [] };
            } else {
                const values = usedRange.values || [];
                const rowCount = usedRange.rowCount || values.length;
                const columnCount = usedRange.columnCount || (values[0] ? values[0].length : 0);
                const headerNames = values.length ? values[0].map(h => (h === undefined || h === null) ? "" : String(h)) : [];

                const headerIndexMap = {};
                headerNames.forEach((h, i) => { headerIndexMap[h] = i; });

                const idColIdx = (identifierColumn && typeof identifierColumn === 'string')
                    ? headerNames.findIndex(h => normalize(h) === normalize(identifierColumn))
                    : -1;

                let data;
                if (idColIdx !== -1) {
                    data = {}; 
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r] || [];
                        const rowObj = {};
                        for (let c = 0; c < columnCount; c++) {
                            const headerKey = headerNames[c] !== "" ? headerNames[c] : `Column${c}`;
                            rowObj[headerKey] = (row[c] !== undefined) ? row[c] : null;
                        }
                        const rawKey = row[idColIdx];
                        const key = (rawKey === undefined || rawKey === null) ? `Row${r}` : String(rawKey);
                        if (!data[key]) data[key] = [];
                        data[key].push(rowObj);
                    }
                } else {
                    data = [];
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r] || [];
                        const rowObj = {};
                        for (let c = 0; c < columnCount; c++) {
                            const headerKey = headerNames[c] !== "" ? headerNames[c] : `Column${c}`;
                            rowObj[headerKey] = (row[c] !== undefined) ? row[c] : null;
                        }
                        data.push(rowObj);
                    }
                }

                if (identifierRow !== null && identifierRow !== undefined) {
                    if (idColIdx !== -1) {
                        const lookupKey = String(identifierRow);
                        data = data[lookupKey] || [];
                    } else {
                        const lookupVal = identifierRow;
                        data = (Array.isArray(data) ? data : []).filter(rowObj =>
                            Object.values(rowObj).some(v => v == lookupVal)
                        );
                    }
                }
                payload = { rowCount, columnCount, address: usedRange.address, headers: headerIndexMap, data };
            }
        });

        console.log("ExcelAPI.loadSheet: finished - success");
        return payload;
    } catch (error) {
        console.log("ExcelAPI.loadSheet: finished - error", error?.message);
        return { success: false, message: error?.message || String(error) };
    }
}

/**
 * Gets the currently selected range in the active worksheet.
 * ... (unchanged)
 */
export async function getSelectedRange(arg1, arg2 = null, options = {}) {
    const isCallbackMode = typeof arg1 === 'function';
    const callback = isCallbackMode ? arg1 : null;
    const COLUMN_ALIASES = isCallbackMode ? arg2 : arg1;

    try {
        const result = await Excel.run(async (context) => {
            console.log("ExcelAPI.getSelectedRange: start");
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const workbookSel = context.workbook.getSelectedRange();
            workbookSel.load(["address", "rowIndex"]);
            await context.sync();

            const payload = await loadRange(context, worksheet, workbookSel, COLUMN_ALIASES);
            if (!payload || !payload.success) {
                if (callback) {
                    try { callback({}); } catch (e) { }
                    return { remove: async () => { } };
                }
                return { success: false, message: payload?.message || "Failed to load range" };
            }

            const normalized = Object.assign({}, payload, {
                singleRow: payload.data || null,
                rows: payload.data ? [payload.data] : []
            });

            if (callback) {
                try {
                    callback({
                        address: payload.address,
                        data: payload.data || {},
                        values: payload.values || [],
                        rowCount: payload.rowCount || 1,
                        allRows: payload.allRows || [],
                        hiddenRowCount: payload.hiddenRowCount || 0
                    });
                } catch (err) {
                    console.error("getSelectedRange callback error:", err);
                }
                return { remove: async () => { } };
            } else {
                console.log("ExcelAPI.getSelectedRange: payload succeeded");
                return normalized;
            }
        });

        return result;
    } catch (error) {
        console.error("ExcelAPI.getSelectedRange: error", error);
        if (isCallbackMode && typeof callback === 'function') {
            try { callback({}); } catch (_) { }
            return { remove: async () => { } };
        }
        return { success: false, message: error && error.message ? error.message : String(error) };
    }
}

/**
 * Registers a worksheet change handler.
 * ... (unchanged)
 */
export async function onChanged(callback, sheet, identifierColumn = null, COLUMN_ALIASES = null, otherValues = null) {
    console.log("ExcelAPI.onChanged: start", { sheet, identifierColumn });
    if (sheet !== undefined && sheet !== null && typeof sheet !== 'string') {
        throw new Error("Parameter 'sheet' must be a string when provided.");
    }
    if (typeof callback !== 'function') {
        throw new Error("Parameter 'callback' is required and must be a function.");
    }

    let eventHandler;
    await Excel.run(async (context) => {
        let worksheet;
        if (sheet && typeof sheet === 'string' && sheet.trim() !== '') {
            worksheet = context.workbook.worksheets.getItemOrNullObject(sheet);
            worksheet.load("isNullObject");
            await context.sync();
            if (worksheet.isNullObject) {
                throw new Error(`Worksheet "${sheet}" not found.`);
            }
        } else {
            worksheet = context.workbook.worksheets.getActiveWorksheet();
        }

        eventHandler = worksheet.onChanged.add(async (eventArgs) => {
            try {
                console.log("ExcelAPI.onChanged: change event fired", { sheet, address: eventArgs && eventArgs.address });
                if (eventArgs && eventArgs.source) {
                    const src = eventArgs.source;
                    const isLocalString = (typeof src === 'string' && src.toLowerCase() === 'local');
                    const isLocalEnum = (typeof Excel !== 'undefined' && Excel.EventSource && src === Excel.EventSource.local);
                    if (!isLocalString && !isLocalEnum) {
                        return;
                    }
                }
                const changedRange = worksheet.getRange(eventArgs.address);
                changedRange.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
                const usedRange = worksheet.getUsedRangeOrNullObject();
                usedRange.load(["values", "rowIndex", "columnIndex", "columnCount", "rowCount", "isNullObject"]);
                await context.sync();

                if (usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
                    // Raw changes (no headers)
                    const rawChanges = [];
                    for (let r = 0; r < (changedRange.rowCount || 0); r++) {
                        for (let c = 0; c < (changedRange.columnCount || 0); c++) {
                            const absRow = changedRange.rowIndex + r;
                            const absCol = changedRange.columnIndex + c;
                            const addr = `${colIndexToLetter(absCol)}${absRow + 1}`;
                            rawChanges.push({
                                address: addr,
                                rowIndex: absRow,
                                colIndex: absCol,
                                header: null,
                                value: (changedRange.values && changedRange.values[r] ? changedRange.values[r][c] : null),
                                identifier: null
                            });
                        }
                    }
                    try { callback({ address: changedRange.address, changes: rawChanges }); } catch (_) { }
                    return;
                }

                const headerRow = usedRange.values[0].map(h => (h === undefined || h === null) ? "" : String(h));
                const usedStartRow = typeof usedRange.rowIndex === "number" ? usedRange.rowIndex : 0;
                const usedStartCol = typeof usedRange.columnIndex === "number" ? usedRange.columnIndex : 0;

                let idColIdx = -1;
                if (identifierColumn && typeof identifierColumn === 'string') {
                    idColIdx = getCanonicalColIdx(headerRow, identifierColumn, canonicalHeaderMap);
                }

                const changes = [];
                for (let r = 0; r < (changedRange.rowCount || 0); r++) {
                    for (let c = 0; c < (changedRange.columnCount || 0); c++) {
                        const relColIndex = (changedRange.columnIndex + c) - usedStartCol;
                        const relRowIndex = (changedRange.rowIndex + r) - usedStartRow;
                        const absRow = changedRange.rowIndex + r;
                        const absCol = changedRange.columnIndex + c;
                        const headerName = (relColIndex >= 0 && relColIndex < headerRow.length) ? headerRow[relColIndex] : `Column${usedStartCol + relColIndex}`;

                        let identifierValue = null;
                        if (idColIdx !== -1) {
                            const idRel = idColIdx;
                            const dataRowIndex = (changedRange.rowIndex + r) - usedStartRow;
                            if (dataRowIndex >= 1 && dataRowIndex < usedRange.values.length) {
                                const dataRow = usedRange.values[dataRowIndex] || [];
                                identifierValue = dataRow[idRel] !== undefined ? dataRow[idRel] : null;
                            } else {
                                if ((changedRange.columnIndex <= usedStartCol + idColIdx) && (usedStartCol + idColIdx < changedRange.columnIndex + changedRange.columnCount)) {
                                    const idOffset = (usedStartCol + idColIdx) - changedRange.columnIndex;
                                    identifierValue = (changedRange.values && changedRange.values[r] ? changedRange.values[r][idOffset] : null);
                                } else {
                                    identifierValue = null;
                                }
                            }
                        }

                        const addr = `${colIndexToLetter(absCol)}${absRow + 1}`;
                        const cellVal = (changedRange.values && changedRange.values[r] ? changedRange.values[r][c] : null);
                        changes.push({
                            address: addr,
                            rowIndex: absRow,
                            colIndex: absCol,
                            header: headerName,
                            value: cellVal,
                            identifier: identifierValue
                        });
                    }
                }

                let filteredChanges = changes;
                if (idColIdx !== -1) {
                    const absIdCol = usedStartCol + idColIdx;
                    filteredChanges = changes.filter(ch => ch.colIndex === absIdCol);
                }

                if (filteredChanges.length === 0) {
                    return;
                }

                if (idColIdx !== -1 && Array.isArray(otherValues) && otherValues.length > 0) {
                    const resolvedCols = otherValues.map(h => {
                        if (!h || typeof h !== 'string') return { name: h, relIdx: -1, absIdx: -1 };
                        const relIdx = getCanonicalColIdx(headerRow, h, canonicalHeaderMap);
                        return { name: h, relIdx, absIdx: (relIdx !== -1 ? usedStartCol + relIdx : -1) };
                    });

                    filteredChanges.forEach(ch => {
                        const rOffset = ch.rowIndex - changedRange.rowIndex;
                        const dataRowIndex = ch.rowIndex - usedStartRow;
                        const map = {};
                        for (const colInfo of resolvedCols) {
                            let val = null;
                            if (colInfo.relIdx !== -1) {
                                if (dataRowIndex >= 1 && dataRowIndex < usedRange.values.length) {
                                    const dataRow = usedRange.values[dataRowIndex] || [];
                                    val = (dataRow[colInfo.relIdx] !== undefined) ? dataRow[colInfo.relIdx] : null;
                                } else {
                                    if (colInfo.absIdx !== -1 &&
                                        changedRange.columnIndex <= colInfo.absIdx &&
                                        colInfo.absIdx < changedRange.columnIndex + changedRange.columnCount &&
                                        rOffset >= 0 && rOffset < (changedRange.values ? changedRange.values.length : 0)) {
                                        const colOffset = colInfo.absIdx - changedRange.columnIndex;
                                        val = (changedRange.values && changedRange.values[rOffset] ? changedRange.values[rOffset][colOffset] : null);
                                    } else {
                                        val = null;
                                    }
                                }
                            } else {
                                val = null;
                            }
                            map[colInfo.name] = val;
                        }
                        ch.otherValues = map;
                    });
                 }

                try { callback({ address: changedRange.address, changes: filteredChanges }); } catch (err) { console.error("onChanged callback error:", err); }
            } catch (err) {
                console.error("onChanged handler error:", err);
            }
        });

        await context.sync();
    });

    console.log("ExcelAPI.onChanged: finished - handler registered");
    return {
        remove: async () => {
            if (eventHandler) {
                await Excel.run(async (context) => {
                    eventHandler.remove();
                    await context.sync();
                });
            }
        }
    };
}