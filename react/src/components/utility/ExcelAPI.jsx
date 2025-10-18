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
    console.log("editRow called with:", { sheet, columnId, rowId, newData });
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
            // Debug: show headers, requested ID column and canonical header map
            console.log("ExcelAPI:debug editRow - headers:", headers);
            console.log("ExcelAPI:debug editRow - requested columnId:", columnId);
            console.log("ExcelAPI:debug editRow - canonicalHeaderMap:", canonicalHeaderMap);

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
                // Debug: show colName being resolved and canonical map
                console.log("ExcelAPI:debug editRow - resolving colName:", colName, { canonicalHeaderMap });
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
        return { success: true, message: "Row updated successfully." };
    } catch (error) {
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
    console.log("insertRow called with:", { sheet, newData });
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        console.log("insertRow: Invalid sheet parameter", sheet);
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }
    try {
        let changeVerified = false;
        let errorMsg = "";
        await Excel.run(async (context) => {
            console.log("insertRow: Getting worksheet", sheet);
            const worksheet = context.workbook.worksheets.getItem(sheet);
            const usedRange = worksheet.getUsedRange();
            usedRange.load(["values", "columnCount", "rowCount"]);
            await context.sync();

            console.log("insertRow: Loaded usedRange", usedRange.values);
            const headers = usedRange.values[0];
            // Debug: show headers and canonical header map for insert
            console.log("ExcelAPI:debug insertRow - headers:", headers);
            console.log("ExcelAPI:debug insertRow - canonicalHeaderMap:", canonicalHeaderMap);

            const newRow = Array(usedRange.columnCount).fill("");
            for (const [colName, value] of Object.entries(newData)) {
                // Debug: show colName being resolved and canonical map
                console.log("ExcelAPI:debug insertRow - resolving colName:", colName, { canonicalHeaderMap });
                // Resolve column by canonical mapping
                const colIdx = getCanonicalColIdx(headers, colName, canonicalHeaderMap);
                console.log(`insertRow: Setting value for column '${colName}' at index ${colIdx}:`, value);
                 if (colIdx !== -1) {
                     newRow[colIdx] = value;
                 }
             }
             console.log("insertRow: New row to insert", newRow);
             worksheet.getRangeByIndexes(usedRange.rowCount, 0, 1, usedRange.columnCount).values = [newRow];
             await context.sync();

             // Double check
             const verifyRange = worksheet.getRangeByIndexes(usedRange.rowCount, 0, 1, usedRange.columnCount);
             verifyRange.load("values");
             await context.sync();
             const verifyValues = verifyRange.values[0];
             console.log("insertRow: Verify inserted row values", verifyValues);
             changeVerified = Object.entries(newData).every(([colName, value]) => {
                 const colIdx = getCanonicalColIdx(headers, colName, canonicalHeaderMap);
                 return colIdx !== -1 && verifyValues[colIdx] == value;
             });
             if (!changeVerified) {
                 errorMsg = "Double check failed: Row was not inserted as expected.";
                 console.log("insertRow: Double check failed", errorMsg);
             } else {
                 console.log("insertRow: Row inserted and verified successfully");
             }
         });
         if (!changeVerified) {
             return { success: false, message: errorMsg };
         }
        return { success: true, message: "Row inserted successfully." };
    } catch (error) {
        console.log("insertRow: Exception occurred", error);
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
    console.log("deleteRow called with:", { sheet, columnId, rowId });
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
            // Debug: show headers, requested ID column and canonical header map for delete
            console.log("ExcelAPI:debug deleteRow - headers:", headers);
            console.log("ExcelAPI:debug deleteRow - requested columnId:", columnId);
            console.log("ExcelAPI:debug deleteRow - canonicalHeaderMap:", canonicalHeaderMap);

            const idColIdx = getCanonicalColIdx(headers, columnId, canonicalHeaderMap);
            if (idColIdx === -1) {
                throw new Error(`ID column "${columnId}" not found.`);
            }

            // Find row index by ID
            let rowIdx = -1;
            for (let i = 1; i < usedRange.rowCount; i++) {
                // Debug: show candidate cell values while searching for ID
                // (will produce more output if you have many rows)
                if (i <= 5) console.log(`ExcelAPI:debug deleteRow - checking row ${i} value:`, usedRange.values[i][idColIdx]);
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
        return { success: true, message: "Row deleted successfully." };
    } catch (error) {
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

    if (typeof Excel === "undefined") {
        throw new Error("Excel API not available.");
    }
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
Example usage for editRow:
editRow("Students", "StudentID", 12345, { Name: "John Doe", Status: "Active" })
  .then(result => console.log(result));

Example usage for insertRow:
insertRow("Students", { StudentID: 67890, Name: "Jane Smith", Status: "Inactive" })
  .then(result => console.log(result));

Example usage for deleteRow:
deleteRow("Students", "StudentID", 12345)
  .then(result => console.log(result));
*/
