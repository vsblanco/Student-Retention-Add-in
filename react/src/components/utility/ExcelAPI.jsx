import { getCanonicalColIdx, canonicalHeaderMap } from './CanonicalMap.jsx';

// [2025-10-23] Version 2.0
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
 *
 * @param {string} sheet The name of the worksheet to edit.
 * @param {string} columnId The name of the column that contains the unique IDs for each row.
 * @param {string | number} rowId The unique ID of the row you want to edit.
 * @param {object} newData An object where keys are the column headers to update and values are the new values for those cells.
 * @returns {Promise<{success: boolean, message: string}>}
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
 *
 * @param {string} sheet The name of the worksheet to insert into.
 * @param {object} newData An object where keys are the column headers and values are the values for those cells.
 * @returns {Promise<{success: boolean, message: string}>}
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
 *
 * @param {string} sheet The name of the worksheet to edit.
 * @param {string} columnId The name of the column that contains the unique IDs for each row.
 * @param {string | number} rowId The unique ID of the row you want to delete.
 * @returns {Promise<{success: boolean, message: string}>}
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
 * 
 * @param {function} callback - Function to run when selection changes.
 * @returns {Promise<{remove: function}>} - Returns an object with a remove() function to unregister the event.
 */
export async function onSelectionChanged(callback, COLUMN_ALIASES = null) {
    console.log("ExcelAPI.onSelectionChanged: start");
    let eventHandler;
    await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        // register handler that delegates work to loadRange
        eventHandler = worksheet.onSelectionChanged.add(async (eventArgs) => {
            try {
                // Delegate selection parsing to loadRange
                const payload = await loadRange(context, worksheet, eventArgs.address, COLUMN_ALIASES);
                if (!payload || !payload.success) {
                    callback({});
                    return;
                }
                // Keep same callback shape as before
                callback({
                    address: payload.address,
                    data: payload.data || {},
                    values: payload.values || []
                });
            } catch (err) {
                console.error("onSelectionChanged callback error:", err);
            }
        });
        await context.sync();
    });
    console.log("ExcelAPI.onSelectionChanged: finished - handler registered");
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

// New: reusable loader for a given range address within an existing Excel.run context.
// This version mirrors the prior onSelectionChanged behavior: it reads the entire usedRange header row,
// then reads the full row (across used columns) for the selected row index and builds a header->value map.
export async function loadRange(context, worksheet, rangeAddress, COLUMN_ALIASES = null) {
    // Load selection rowIndex (we only need the selected row index for this address)
    const selRange = worksheet.getRange(rangeAddress);
    selRange.load(["address", "rowIndex"]);
    // Load usedRange headers and position
    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load(["values", "formulas", "rowIndex", "columnIndex", "columnCount", "isNullObject"]);
    await context.sync();

    if (!usedRange || usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
        return { success: false, message: "No used range / headers found.", address: rangeAddress };
    }

    // Normalize raw header strings
    const rawHeaders = usedRange.values[0].map(h => (h === undefined || h === null) ? "" : String(h));
    const usedColIndex = (typeof usedRange.columnIndex === "number") ? usedRange.columnIndex : 0;
    const usedColCount = usedRange.columnCount || (rawHeaders.length);

    // Ensure selRange.rowIndex is available (already loaded earlier)
    const selectedRowIndex = selRange.rowIndex;

    // Read the full row across the used columns
    const rowRange = worksheet.getRangeByIndexes(selectedRowIndex, usedColIndex, 1, usedColCount);
    rowRange.load("values");
    await context.sync();

    const rowValues = (rowRange.values && rowRange.values[0]) ? rowRange.values[0] : [];
    // Compute relative row index into usedRange.formulas (if available)
    const usedRangeStartRow = (typeof usedRange.rowIndex === "number") ? usedRange.rowIndex : 0;
    const relativeRowIdx = selectedRowIndex - usedRangeStartRow;
    const formulasRow = (usedRange.formulas && usedRange.formulas[relativeRowIdx]) ? usedRange.formulas[relativeRowIdx] : [];

    // Build header -> value map for the entire row, using COLUMN_ALIASES if provided
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

    // Build headers index map
    const headerIndexMap = {};
    rawHeaders.forEach((h, i) => { headerIndexMap[h] = i; });

    return {
        success: true,
        address: selRange.address,
        rowIndex: selectedRowIndex,
        startCol: usedColIndex,
        columnCount: usedColCount,
        headers: headerIndexMap,
        values: rowValues,
        formulas: formulasRow,
        data: rowObj
    };
}

/**
 * Checks if a row exists in the worksheet by unique ID.
 *
 * @param {string} sheet The name of the worksheet to check.
 * @param {string} columnId The name of the column that contains the unique IDs for each row.
 * @param {string | number} rowId The unique ID of the row you want to check.
 * @returns {Promise<boolean>} Promise resolving to true if the row exists, false otherwise.
 */
export async function checkRow(sheet, columnId, rowId) {
     // Quick validation
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
                 // If ID column not found, row cannot exist
                 exists = false;
                 return;
             }

             // Scan rows (skip header row at index 0)
             for (let i = 1; i < usedRange.rowCount; i++) {
                 if (usedRange.values[i][idColIdx] == rowId) { // loose equality to match numbers/strings
                    exists = true;
                    break;
                 }
             }
         });
         console.log("ExcelAPI.checkRow: finished - result", exists);
         return exists;
     } catch (error) {
         console.error("ExcelAPI.checkRow: finished - error", error && error.message ? error.message : error);
         // Surface errors as rejected promise
         throw new Error(error.message || String(error));
     }
}

/**
 * Highlights a row in the active worksheet.
 *
 * @param {number} rowIndex The index of the row to highlight (0-based).
 * @param {number} startCol The index of the reference column in the range (0-based).
 *                         Highlighting will extend to the left of this column.
 * @param {number} colCount The number of columns to include in the highlight (extends left).
 * @param {string} [color='yellow'] The color to use for highlighting (in CSS color format).
 */
export async function highlightRow(rowIndex, startCol, colCount, color = 'yellow') {
  // Quick guard: Excel runtime must exist and parameters must be numbers
  if (typeof window.Excel === "undefined") return;
  if (typeof rowIndex !== 'number' || typeof startCol !== 'number' || typeof colCount !== 'number') return;
  if (colCount <= 0) return;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Compute leftmost column so the range extends to the left of startCol.
      // Example: startCol=4, colCount=3 -> leftStartCol = 4 - (3 - 1) = 2  (columns 2,3,4)
      let leftStartCol = startCol - (colCount - 1);
      if (leftStartCol < 0) {
        // clamp to sheet left edge and adjust count so we don't request negative index
        leftStartCol = 0;
      }
      // Adjust actual column count in case leftStartCol was clamped
      const actualColCount = (startCol - leftStartCol) + 1;
      if (actualColCount <= 0) return;

      const highlightRange = sheet.getRangeByIndexes(rowIndex, leftStartCol, 1, actualColCount);
      highlightRange.format.fill.color = color;
      await context.sync();
    });
  } catch (_) {
    // swallow errors intentionally to avoid breaking callers
  }
}

/**
 * Loads a worksheet's used range and returns headers, values and dimensions.
 *
 * @param {string} sheet The name of the worksheet to load.
 * @returns {Promise<{success: boolean, message: string, data: {headers: any[], values: any[][], rowCount: number, columnCount: number} | null}>}
 */
export async function loadSheet(sheet, identifierColumn = null, identifierRow = null) {
    console.log("ExcelAPI.loadSheet: start", { sheet }, { identifierColumn });
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        console.log("ExcelAPI.loadSheet: finished - invalid parameters");
        // removed parent `data` per request
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
                // empty sheet (no used range)
                payload = { rowCount: 0, columnCount: 0, headers: {}, data: [] };
            } else {
                const values = usedRange.values || [];
                const rowCount = usedRange.rowCount || values.length;
                const columnCount = usedRange.columnCount || (values[0] ? values[0].length : 0);
                // Normalize headers to strings (use empty string for undefined/null)
                const headerNames = values.length ? values[0].map(h => (h === undefined || h === null) ? "" : String(h)) : [];

                // Build header -> index map (expose this under `headers` in the result)
                const headerIndexMap = {};
                headerNames.forEach((h, i) => {
                    headerIndexMap[h] = i;
                });

                // Determine identifier column index by direct presence in headerNames (no canonical matching)
                const idColIdx = (identifierColumn && typeof identifierColumn === 'string')
                    ? headerNames.findIndex(h => normalize(h) === normalize(identifierColumn))
                    : -1;

                // Build `data`:
                // - if idColIdx !== -1 => produce an object keyed by identifier value, each value is ALWAYS an array of rows
                // - otherwise => produce an array of row objects (legacy behavior)
                let data;
                if (idColIdx !== -1) {
                    data = {}; // keyed by identifier value -> array of row objects
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r] || [];
                        const rowObj = {};
                        for (let c = 0; c < columnCount; c++) {
                            const headerKey = headerNames[c] !== "" ? headerNames[c] : `Column${c}`;
                            rowObj[headerKey] = (row[c] !== undefined) ? row[c] : null;
                        }
                        // derive key from the identifier column value; coerce to string to be safe
                        const rawKey = row[idColIdx];
                        const key = (rawKey === undefined || rawKey === null) ? `Row${r}` : String(rawKey);
                        if (!data[key]) data[key] = [];
                        data[key].push(rowObj);
                    }
                } else {
                    // fallback: array of row objects
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

                // If identifierRow was provided, filter the result accordingly:
                if (identifierRow !== null && identifierRow !== undefined) {
                    if (idColIdx !== -1) {
                        // identifierColumn present: return only the array for that identifier key (or empty array)
                        const lookupKey = String(identifierRow);
                        data = data[lookupKey] || [];
                    } else {
                        // identifierColumn not present: treat identifierRow as a filter value and
                        // return only rows where any cell loosely equals identifierRow
                        const lookupVal = identifierRow;
                        data = (Array.isArray(data) ? data : []).filter(rowObj =>
                            Object.values(rowObj).some(v => v == lookupVal)
                        );
                    }
                }

                // Note: `headers` here is the index map (name -> column index)
                payload = { rowCount, columnCount, address: usedRange.address, headers: headerIndexMap, data };
            }
        });

        console.log("ExcelAPI.loadSheet: finished - success");
        // Return the payload (data object) directly so callers receive the sheet payload
        return payload;
    } catch (error) {
        console.log("ExcelAPI.loadSheet: finished - error", error?.message);
        // removed parent `data` per request
        return { success: false, message: error?.message || String(error) };
    }
}

/**
 * Gets the currently selected range in the active worksheet.
 *
 * Returns metadata and canonicalized row objects for the current selection.
 * @param {function} callback - Function to run when selection changes.
 * @param {object} COLUMN_ALIASES - Optional mapping of canonical column names to alias names.
 * @param {object} options - Optional flags for the operation (currently unused).
 * @returns {Promise<{success: boolean, message: string, address: string, startRow: number, startCol: number, rowCount: number, columnCount: number, headers: object, values: any[][], formulas: any[][], rows: object[], singleRow: object | null}>}
 */
export async function getSelectedRange(arg1, arg2 = null, options = {}) {
    // Supports two call styles:
    // 1) callback mode: getSelectedRange(callback, COLUMN_ALIASES)
    // 2) promise mode:  getSelectedRange(COLUMN_ALIASES)
    const isCallbackMode = typeof arg1 === 'function';
    const callback = isCallbackMode ? arg1 : null;
    const COLUMN_ALIASES = isCallbackMode ? arg2 : arg1;

    try {
        // Use Excel.run and return its result so callers can await the payload
        const result = await Excel.run(async (context) => {
            console.log("ExcelAPI.getSelectedRange: start");
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const workbookSel = context.workbook.getSelectedRange();
            workbookSel.load(["address", "rowIndex"]);
            await context.sync();

            const payload = await loadRange(context, worksheet, workbookSel.address, COLUMN_ALIASES);
            if (!payload || !payload.success) {
                if (callback) {
                    try { callback({}); } catch (e) { /* swallow callback errors */ }
                    return { remove: async () => { /* noop for one-time selection */ } };
                }
                return { success: false, message: payload?.message || "Failed to load range" };
            }

            // Normalize payload for promise-mode consumers
            const normalized = Object.assign({}, payload, {
                singleRow: payload.data || null,
                rows: payload.data ? [payload.data] : []
            });

            if (callback) {
                try {
                    callback({
                        address: payload.address,
                        data: payload.data || {},
                        values: payload.values || []
                    });
                } catch (err) {
                    console.error("getSelectedRange callback error:", err);
                }
                // mirror onSelectionChanged shape (remove() available)
                return { remove: async () => { /* noop - one-time selection */ } };
            } else {
                console.log("ExcelAPI.getSelectedRange: payload succeeded");
                return normalized;
            }
        });

        return result;
    } catch (error) {
        console.error("ExcelAPI.getSelectedRange: error", error);
        if (isCallbackMode && typeof callback === 'function') {
            try { callback({}); } catch (_) { /* ignore */ }
            return { remove: async () => { /* noop */ } };
        }
        return { success: false, message: error && error.message ? error.message : String(error) };
    }
}

/**
 * Registers a worksheet change handler.
 *
 * @param {string} sheet - Worksheet name to observe (required).
 * @param {string|null} identifierColumn - Column (name) that serves as row identifier (optional).
 * @param {function} callback - function(payload) called when changes occur. payload = { address, changes: [{ address, rowIndex, colIndex, header, value, identifier }] }
 * @param {object|null} COLUMN_ALIASES - optional alias map for canonical header resolution.
 * @returns {Promise<{remove: function}>} - object with remove() to unregister.
 */
export async function onChanged(callback, sheet, identifierColumn = null, COLUMN_ALIASES = null, otherValues = null) {
    console.log("ExcelAPI.onChanged: start", { sheet, identifierColumn });
    // Allow omitted sheet: if provided it must be a string; otherwise we'll use the active worksheet.
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
            // use the named worksheet (preserve previous behavior)
            worksheet = context.workbook.worksheets.getItemOrNullObject(sheet);
            worksheet.load("isNullObject");
            await context.sync();

            if (worksheet.isNullObject) {
                throw new Error(`Worksheet "${sheet}" not found.`);
            }
        } else {
            // no sheet provided -> use the currently active worksheet
            worksheet = context.workbook.worksheets.getActiveWorksheet();
        }

        // Register handler which uses the same context and worksheet captures
        eventHandler = worksheet.onChanged.add(async (eventArgs) => {
            try {
               // Log that a change event fired (include sheet and event address)
                console.log("ExcelAPI.onChanged: change event fired", { sheet, address: eventArgs && eventArgs.address });
                // Load changed range and used range within this captured context
                const changedRange = worksheet.getRange(eventArgs.address);
                changedRange.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
                const usedRange = worksheet.getUsedRangeOrNullObject();
                usedRange.load(["values", "rowIndex", "columnIndex", "columnCount", "rowCount", "isNullObject"]);
                await context.sync();

                if (usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
                    // No headers / used range — still report raw changed cells
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
                    console.log("ExcelAPI.onChanged: computed rawChanges", { sheet, address: changedRange.address, changes: rawChanges });
                    try { callback({ address: changedRange.address, changes: rawChanges }); } catch (_) { /* swallow */ }
                    return;
                }

                const headerRow = usedRange.values[0].map(h => (h === undefined || h === null) ? "" : String(h));
                const usedStartRow = typeof usedRange.rowIndex === "number" ? usedRange.rowIndex : 0;
                const usedStartCol = typeof usedRange.columnIndex === "number" ? usedRange.columnIndex : 0;

                // Determine identifier column index (relative to usedRange header indices) using canonical mapping
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

                        // identifier value: if idColIdx is present and the usedRange covers that row, extract it
                        let identifierValue = null;
                        if (idColIdx !== -1) {
                            const idRel = idColIdx; // idColIdx is already relative to headerRow indices
                            // Use absolute row index to compute data row inside usedRange (header is at index 0)
                            const dataRowIndex = (changedRange.rowIndex + r) - usedStartRow;
                            // data rows in usedRange.values start at index 1 (0 == header). Ensure we are inside bounds.
                            if (dataRowIndex >= 1 && dataRowIndex < usedRange.values.length) {
                                const dataRow = usedRange.values[dataRowIndex] || [];
                                identifierValue = dataRow[idRel] !== undefined ? dataRow[idRel] : null;
                            } else {
                                // If usedRange.values doesn't include this row, try to read changedRange value in id column if it's within changedRange
                                if ((changedRange.columnIndex <= usedStartCol + idColIdx) && (usedStartCol + idColIdx < changedRange.columnIndex + changedRange.columnCount)) {
                                    const idOffset = (usedStartCol + idColIdx) - changedRange.columnIndex;
                                    identifierValue = (changedRange.values && changedRange.values[r] ? changedRange.values[r][idOffset] : null);
                                } else {
                                    // fallback: null
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

                // If an identifier column was resolved, only keep changes that occurred in that column.
                let filteredChanges = changes;
                if (idColIdx !== -1) {
                    const absIdCol = usedStartCol + idColIdx;
                    filteredChanges = changes.filter(ch => ch.colIndex === absIdCol);
                    console.log("ExcelAPI.onChanged: identifier column present, filtered changes", { identifierColumn, absIdCol, before: changes.length, after: filteredChanges.length });
                }

                // If there are no relevant changes after filtering (when identifier column exists), do not invoke callback.
                if (filteredChanges.length === 0) {
                    // Log why callback is not sent when identifier column filtering removed all changes
                    if (idColIdx !== -1) {
                        const absIdCol = usedStartCol + idColIdx;
                        console.log("ExcelAPI.onChanged: callback not sent because change was not under identifier column", { identifierColumn, absIdCol, totalDetectedChanges: changes.length });
                    } else {
                        console.log("ExcelAPI.onChanged: callback not sent because no relevant changes were detected");
                    }
                    // nothing to report
                    return;
                }

                // NEW: if otherValues was provided as an array of header names, resolve and attach their values for the affected row(s)
                if (idColIdx !== -1 && Array.isArray(otherValues) && otherValues.length > 0) {
                    // Resolve requested headers to relative column indices (relative to headerRow)
                    const resolvedCols = otherValues.map(h => {
                        if (!h || typeof h !== 'string') return { name: h, relIdx: -1, absIdx: -1 };
                        const relIdx = getCanonicalColIdx(headerRow, h, canonicalHeaderMap);
                        return { name: h, relIdx, absIdx: (relIdx !== -1 ? usedStartCol + relIdx : -1) };
                    });

                    // For each filtered change, compute the otherValues mapping
                    filteredChanges.forEach(ch => {
                        const rOffset = ch.rowIndex - changedRange.rowIndex; // row index inside changedRange values
                        // compute index inside usedRange.values (header is at index 0, data rows start at 1)
                        const dataRowIndex = ch.rowIndex - usedStartRow;
                        const map = {};
                        for (const colInfo of resolvedCols) {
                            let val = null;
                            if (colInfo.relIdx !== -1) {
                                // Prefer usedRange.values when the row is present in usedRange
                                if (dataRowIndex >= 1 && dataRowIndex < usedRange.values.length) {
                                    const dataRow = usedRange.values[dataRowIndex] || [];
                                    val = (dataRow[colInfo.relIdx] !== undefined) ? dataRow[colInfo.relIdx] : null;
                                } else {
                                    // fallback to changedRange if the desired column sits inside changedRange
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

                // Log computed changes before invoking callback
                console.log("ExcelAPI.onChanged: computed changes", { sheet, address: changedRange.address, changes: filteredChanges });
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

/*
Example usages (updated to new return shape):

// Load sheet and inspect headers/values
(async () => {
    const res = await loadSheet("Students");
    if (res.success) {
        console.log("Sheet headers (index map):", res.headers);
        console.log("Row count:", res.rowCount);
        console.log("Column count:", res.columnCount);
        console.log("First few rows:", res.data.slice(0, 5));
    } else {
        console.error("loadSheet failed:", res.message);
    }
})();

// Edit a row by ID
editRow("Students", "StudentID", 12345, { Name: "John Doe", Status: "Active" })
  .then(result => console.log("editRow:", result));

// Insert a new row (values resolved by canonical headers)
insertRow("Students", { StudentID: 67890, Name: "Jane Smith", Status: "Inactive" })
  .then(result => console.log("insertRow:", result));

// Delete a row by ID
deleteRow("Students", "StudentID", 12345)
  .then(result => console.log("deleteRow:", result));

// Check if a row exists
checkRow("Students", "StudentID", 12345)
  .then(exists => console.log("Row exists:", exists))
  .catch(err => console.error("checkRow error:", err));

// Highlight a row (rowIndex is 0-based; pass header row index 0 if needed)
await highlightRow(2, 0, 5, 'lightgreen'); // highlights row index 2 across first 5 columns

// Register a selection-changed handler and remove it later
(async () => {
    const handler = await onSelectionChanged(({ address, values }) => {
        console.log("Selection changed:", address, values);
    });
    // later to remove:
    // await handler.remove();
})();

// Register a change handler (logs changes to console)
const changeHandler = await onChanged("Students", "StudentID", ({ address, changes }) => {
    console.log("Changes detected:", address, changes);
});

// To unregister the change handler:
// await changeHandler.remove();
*/
