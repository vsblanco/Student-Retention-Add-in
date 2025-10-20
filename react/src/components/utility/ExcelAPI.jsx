/**
 * Edits a row in an Excel worksheet by unique ID.
 *
 * @param {string} sheet The name of the worksheet to edit.
 * @param {string} columnId The name of the column that contains the unique IDs for each row.
 * @param {string | number} rowId The unique ID of the row you want to edit.
 * @param {object} newData An object where keys are the column headers to update and values are the new values for those cells.
 * @returns {Promise<{success: boolean, message: string}>}
 */
import { getCanonicalColIdx, canonicalHeaderMap } from './CanonicalMap.jsx';

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
export async function onSelectionChanged(callback) {
    console.log("ExcelAPI.onSelectionChanged: start");
     let eventHandler;
     await Excel.run(async (context) => {
         const worksheet = context.workbook.worksheets.getActiveWorksheet();
         eventHandler = worksheet.onSelectionChanged.add(async (eventArgs) => {
             try {
                 const range = worksheet.getRange(eventArgs.address);
                 range.load("values");
                 await context.sync();
                 callback({
                     address: eventArgs.address,
                     values: range.values,
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

/*
Example usage for editRow:
editRow("Students", "StudentID", 12345, { Name: "John Doe", Status: "Active" })
  .then(result => console.log(result));

Example usage for insertRow:
insertRow("Students", { StudentID: 67890, Name: "Jane Smith", Status: "Inactive" })
  .then(result => console.log(result));

Example usage for deleteRow:
deleteRow("Students", "StudentID", 12345)
  .then(result => console.log(result));

Example usage for checkRow:
checkRow("Students", "StudentID", 12345)
  .then(exists => console.log("Row exists:", exists));
*/
