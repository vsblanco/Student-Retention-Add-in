import { getCanonicalColIdx, canonicalHeaderMap } from './CanonicalMap.jsx';

// [2025-10-23] Version 2.0
export function normalize(val) {
    if (val === undefined || val === null) return "";
    // Convert to string, remove all spaces globally, and convert to lowercase.
    return String(val).replace(/\s/g, '').toLowerCase();
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
        // register handler
        eventHandler = worksheet.onSelectionChanged.add(async (eventArgs) => {
            try {
                // Load selection rowIndex (we only care about the first row of the selection)
                const selRange = worksheet.getRange(eventArgs.address);
                selRange.load(["rowIndex"]);
                // Load usedRange headers and position
                const usedRange = worksheet.getUsedRangeOrNullObject();
                usedRange.load(["values", "columnIndex", "columnCount", "isNullObject"]);
                await context.sync();

                // If no used range / no headers, return empty object
                if (!usedRange || usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
                    callback({});
                    return;
                }

                // Normalize raw header strings
                const rawHeaders = usedRange.values[0].map(h => (h === undefined || h === null) ? "" : String(h));
                const usedColIndex = (typeof usedRange.columnIndex === "number") ? usedRange.columnIndex : 0;
                const usedColCount = usedRange.columnCount || (rawHeaders.length);
                console.log(rawHeaders);

                await context.sync(); // ensure selRange.rowIndex is available

                const selectedRowIndex = selRange.rowIndex;
                // Read the full row across the used columns
                const rowRange = worksheet.getRangeByIndexes(selectedRowIndex, usedColIndex, 1, usedColCount);
                rowRange.load("values");
                await context.sync();

                const rowValues = (rowRange.values && rowRange.values[0]) ? rowRange.values[0] : [];

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
                            console.log (`Mapping header "${headerRaw}" to canonical "${match}"`);
                            headerKey = match;
                        }
                    }

                    if (!headerKey || headerKey === "") {
                        headerKey = `Column${usedColIndex + c}`;
                    }
                    rowObj[headerKey] = (rowValues[c] !== undefined) ? rowValues[c] : null;
                }

                // Provide both raw values array and the canonicalized data object
                callback({
                    address: eventArgs.address,
                    data: rowObj,
                    values: rowValues
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
 * @param {number} startCol The index of the first column in the range to highlight (0-based).
 * @param {number} colCount The number of columns to include in the highlight.
 * @param {string} [color='yellow'] The color to use for highlighting (in CSS color format).
 */
export async function highlightRow(rowIndex, startCol, colCount, color = 'yellow') {
  // Quick guard: Excel runtime must exist and parameters must be numbers
  if (typeof window.Excel === "undefined") return;
  if (typeof rowIndex !== 'number' || typeof startCol !== 'number' || typeof colCount !== 'number') return;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const highlightRange = sheet.getRangeByIndexes(rowIndex, startCol, 1, colCount);
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
*/
