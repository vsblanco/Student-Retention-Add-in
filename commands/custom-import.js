// Timestamp: 2025-09-18 01:08 PM EDT
// Version: 1.2.0
/*
 * This file contains the logic for handling custom, schema-driven imports
 * from a single JSON file that contains both the schema and the data.
 */

/**
 * Main function to handle the custom import process from a self-contained JSON file.
 * @param {object} message The message from the dialog, containing the file data.
 * @param {function} sendMessageToDialog A function to send status messages back to the dialog.
 */
export async function handleCustomImport(message, sendMessageToDialog) {
    try {
        sendMessageToDialog("Starting custom import from JSON file...");

        // 1. Parse and Validate the JSON Schema and Data
        const importFile = parseAndValidateImportFile(message.data, sendMessageToDialog);
        if (!importFile) return; // Validation failed

        const { schema, data: sourceData } = importFile;

        sendMessageToDialog(`Successfully validated schema: "${schema.importName}".`);
        sendMessageToDialog(`Target sheet: "${schema.targetSheet}".`);
        sendMessageToDialog(`Found ${sourceData.length} data rows embedded in the file.`);

        // 2. Prepare for Excel Operations
        await Excel.run(async (context) => {
            // Get or create the target sheet
            let sheet = context.workbook.worksheets.getItemOrNullObject(schema.targetSheet);
            await context.sync();

            if (sheet.isNullObject) {
                sendMessageToDialog(`Sheet "${schema.targetSheet}" not found, creating it...`);
                sheet = context.workbook.worksheets.add(schema.targetSheet);
                const headers = schema.columnMappings.map(m => m.target);
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                await context.sync();
            } else {
                sendMessageToDialog(`Found existing sheet: "${schema.targetSheet}".`);
            }

            // 3. Read Existing Data from the Target Sheet
            const usedRange = sheet.getUsedRange(true);
            usedRange.load("values, rowCount, address");
            await context.sync();

            const sheetHeaders = usedRange.values.length > 0 ? usedRange.values[0].map(h => String(h || '')) : [];
            const sheetKeyColumnIndex = sheetHeaders.indexOf(schema.sheetKeyColumn);

            if (sheetKeyColumnIndex === -1) {
                throw new Error(`The key column "${schema.sheetKeyColumn}" was not found in the target sheet.`);
            }

            const sheetDataMap = new Map();
            if (usedRange.rowCount > 1) {
                for (let i = 1; i < usedRange.rowCount; i++) {
                    const key = usedRange.values[i][sheetKeyColumnIndex];
                    if (key) {
                        sheetDataMap.set(String(key), { rowIndex: i, data: usedRange.values[i] });
                    }
                }
            }
            sendMessageToDialog(`Mapped ${sheetDataMap.size} existing rows from "${schema.targetSheet}".`);

            // 4. Prepare Data for Writing (Updates and New Rows)
            const rowsToUpdate = [];
            const rowsToAdd = [];
            const targetColumnIndexMap = new Map(sheetHeaders.map((header, index) => [header, index]));
            
            // Find the source key from the mappings to access it in the data objects
            const sourceKey = schema.columnMappings.find(m => m.target === schema.sheetKeyColumn)?.source;
            if (!sourceKey) {
                 throw new Error(`Could not find a 'source' mapping for the sheetKeyColumn "${schema.sheetKeyColumn}".`);
            }

            sourceData.forEach(sourceObject => {
                const key = sourceObject[sourceKey];
                if (key === undefined) return; // Skip if the key is not in the object

                const newRowData = new Array(sheetHeaders.length).fill(null);
                schema.columnMappings.forEach(mapping => {
                    const targetIndex = targetColumnIndexMap.get(mapping.target);
                    if (targetIndex !== undefined && sourceObject[mapping.source] !== undefined) {
                        newRowData[targetIndex] = sourceObject[mapping.source];
                    }
                });

                if (sheetDataMap.has(String(key))) {
                    const existingRow = sheetDataMap.get(String(key));
                    rowsToUpdate.push({ rowIndex: existingRow.rowIndex, values: newRowData });
                } else {
                    rowsToAdd.push(newRowData);
                }
            });

            sendMessageToDialog(`Prepared ${rowsToUpdate.length} rows for update and ${rowsToAdd.length} new rows to add.`);

            // 5. Perform Bulk Write Operations
            if (rowsToUpdate.length > 0) {
                sendMessageToDialog("Writing updates to the sheet...");
                for (const row of rowsToUpdate) {
                    const range = sheet.getRangeByIndexes(row.rowIndex, 0, 1, sheetHeaders.length);
                    range.values = [row.values];
                }
            }

            if (rowsToAdd.length > 0) {
                sendMessageToDialog("Adding new rows to the sheet...");
                const startRow = usedRange.rowCount > 0 ? usedRange.rowCount : 1;
                const range = sheet.getRangeByIndexes(startRow, 0, rowsToAdd.length, sheetHeaders.length);
                range.values = rowsToAdd;
            }

            sheet.getUsedRange().format.autofitColumns();
            await context.sync();
        });

        sendMessageToDialog("Custom import completed successfully.", 'complete');

    } catch (error) {
        console.error("Custom Import Error: ", error);
        sendMessageToDialog(`Error: ${error.message}`, 'error');
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug Info: " + JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Parses and validates the provided JSON import file.
 * @param {string} dataUrl The data URL of the JSON file.
 * @param {function} sendMessageToDialog Function to send status messages.
 * @returns {object|null} An object with 'schema' and 'data' properties, or null if validation fails.
 */
function parseAndValidateImportFile(dataUrl, sendMessageToDialog) {
    try {
        const jsonString = atob(dataUrl.split(',')[1]);
        const importFile = JSON.parse(jsonString);

        // --- Schema Validation ---
        if (!importFile.importName || typeof importFile.importName !== 'string') throw new Error("Missing a valid 'importName'.");
        if (!importFile.targetSheet || typeof importFile.targetSheet !== 'string') throw new Error("Missing a valid 'targetSheet'.");
        if (!importFile.sheetKeyColumn || typeof importFile.sheetKeyColumn !== 'string') throw new Error("Missing a valid 'sheetKeyColumn'.");
        if (!Array.isArray(importFile.columnMappings) || importFile.columnMappings.length === 0) throw new Error("Must have a non-empty 'columnMappings' array.");

        for (const mapping of importFile.columnMappings) {
            if (!mapping.source || typeof mapping.source !== 'string' || !mapping.target || typeof mapping.target !== 'string') {
                throw new Error("Each item in 'columnMappings' must have valid 'source' and 'target' properties.");
            }
        }
        
        // --- Data Validation ---
        if (!Array.isArray(importFile.data)) {
            throw new Error("The 'data' property must be an array.");
        }

        const schema = {
            importName: importFile.importName,
            targetSheet: importFile.targetSheet,
            sheetKeyColumn: importFile.sheetKeyColumn,
            columnMappings: importFile.columnMappings,
        };

        return { schema, data: importFile.data };
    } catch (error) {
        sendMessageToDialog(`JSON file validation failed: ${error.message}`, 'error');
        return null;
    }
}
