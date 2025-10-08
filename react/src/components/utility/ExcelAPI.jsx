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
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }

    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet);
            const usedRange = worksheet.getUsedRange();
            usedRange.load(["values", "columnCount", "rowCount"]);
            await context.sync();

            const headers = usedRange.values[0];
            const idColIdx = headers.indexOf(columnId);
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
                const colIdx = headers.indexOf(colName);
                if (colIdx === -1) continue; // skip unknown columns
                worksheet.getCell(rowIdx, colIdx).values = [[value]];
            }
            await context.sync();
        });
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
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet);
            const usedRange = worksheet.getUsedRange();
            usedRange.load(["values", "columnCount", "rowCount"]);
            await context.sync();

            const headers = usedRange.values[0];
            const newRow = Array(usedRange.columnCount).fill("");
            for (const [colName, value] of Object.entries(newData)) {
                const colIdx = headers.indexOf(colName);
                if (colIdx !== -1) {
                    newRow[colIdx] = value;
                }
            }
            worksheet.getRangeByIndexes(usedRange.rowCount, 0, 1, usedRange.columnCount).values = [newRow];
            await context.sync();
        });
        return { success: true, message: "Row inserted successfully." };
    } catch (error) {
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
    if (!sheet || typeof sheet !== "string" || sheet.trim() === "") {
        return { success: false, message: "Parameter 'sheet' is required and must be a non-empty string." };
    }
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(sheet);
            const usedRange = worksheet.getUsedRange();
            usedRange.load(["values", "columnCount", "rowCount"]);
            await context.sync();

            const headers = usedRange.values[0];
            const idColIdx = headers.indexOf(columnId);
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
        });
        return { success: true, message: "Row deleted successfully." };
    } catch (error) {
        return { success: false, message: error.message };
    }
}