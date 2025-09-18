// Timestamp: 2025-09-18 01:24 PM EDT
// Version: 1.3.0
/*
 * This file contains the logic for handling custom, schema-driven imports
 * from a single JSON file that contains both the schema and the data.
 * Version 1.3.0 introduces support for multi-sheet imports by specifying
 * an optional 'targetSheet' on individual column mappings.
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
        if (!importFile) return;

        const { schema, data: sourceData } = importFile;
        sendMessageToDialog(`Successfully validated schema: "${schema.importName}".`);

        // 2. Identify all unique target sheets from the schema
        const allSheetNames = new Set([schema.targetSheet]);
        schema.columnMappings.forEach(m => {
            if (m.targetSheet) allSheetNames.add(m.targetSheet);
        });
        sendMessageToDialog(`Preparing to process ${allSheetNames.size} target sheet(s).`);

        await Excel.run(async (context) => {
            const sheetCache = new Map();
            const sheetDataCache = new Map();

            // 3. Ensure all target sheets exist, creating them if necessary
            for (const sheetName of allSheetNames) {
                let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                await context.sync();

                if (sheet.isNullObject) {
                    sendMessageToDialog(`Sheet "${sheetName}" not found, creating it...`);
                    sheet = context.workbook.worksheets.add(sheetName);
                    // Find all headers for this new sheet
                    const headers = schema.columnMappings
                        .filter(m => (m.targetSheet || schema.targetSheet) === sheetName)
                        .map(m => m.target);
                    // Ensure the key column is always included
                    if (!headers.includes(schema.sheetKeyColumn)) {
                        headers.unshift(schema.sheetKeyColumn);
                    }
                    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                    headerRange.values = [headers];
                    headerRange.format.font.bold = true;
                } else {
                    sendMessageToDialog(`Found existing sheet: "${sheetName}".`);
                }
                sheetCache.set(sheetName, sheet);
            }
            await context.sync();

            // 4. Read existing data from all target sheets into memory
            for (const [sheetName, sheet] of sheetCache.entries()) {
                const usedRange = sheet.getUsedRange(true);
                usedRange.load("values, rowCount");
                await context.sync();

                const values = usedRange.values || [];
                const headers = values.length > 0 ? values[0].map(h => String(h || '')) : [];
                const keyColumnIndex = headers.indexOf(schema.sheetKeyColumn);

                if (keyColumnIndex === -1) {
                    throw new Error(`The key column "${schema.sheetKeyColumn}" was not found in sheet "${sheetName}". Please add it and try again.`);
                }

                const dataMap = new Map();
                for (let i = 1; i < usedRange.rowCount; i++) {
                    const key = values[i][keyColumnIndex];
                    if (key) dataMap.set(String(key), { rowIndex: i, data: values[i] });
                }
                sheetDataCache.set(sheetName, { headers, keyColumnIndex, dataMap, rowCount: usedRange.rowCount });
                sendMessageToDialog(`Mapped ${dataMap.size} existing rows from "${sheetName}".`);
            }

            // 5. Prepare data for writing by grouping updates and additions by sheet
            const writesBySheet = new Map();
            allSheetNames.forEach(name => writesBySheet.set(name, { rowsToUpdate: [], rowsToAdd: [] }));
            const sourceKey = schema.columnMappings.find(m => m.target === schema.sheetKeyColumn)?.source;
            if (!sourceKey) throw new Error(`Could not find a 'source' mapping for the sheetKeyColumn "${schema.sheetKeyColumn}".`);

            sourceData.forEach(sourceObject => {
                const key = sourceObject[sourceKey];
                if (key === undefined) return;

                const valuesBySheet = new Map();
                schema.columnMappings.forEach(mapping => {
                    const sheetName = mapping.targetSheet || schema.targetSheet;
                    if (!valuesBySheet.has(sheetName)) valuesBySheet.set(sheetName, {});
                    if (sourceObject[mapping.source] !== undefined) {
                        valuesBySheet.get(sheetName)[mapping.target] = sourceObject[mapping.source];
                    }
                });

                for (const [sheetName, newValues] of valuesBySheet.entries()) {
                    const sheetInfo = sheetDataCache.get(sheetName);
                    if (!sheetInfo) continue;

                    const newRowData = new Array(sheetInfo.headers.length).fill(null);
                    sheetInfo.headers.forEach((header, index) => {
                        if (newValues[header] !== undefined) newRowData[index] = newValues[header];
                    });

                    if (sheetInfo.dataMap.has(String(key))) {
                        const existingRow = sheetInfo.dataMap.get(String(key));
                        const finalRowData = existingRow.data.map((val, idx) => newRowData[idx] !== null ? newRowData[idx] : val);
                        writesBySheet.get(sheetName).rowsToUpdate.push({ rowIndex: existingRow.rowIndex, values: finalRowData });
                    } else {
                        newRowData[sheetInfo.keyColumnIndex] = key;
                        writesBySheet.get(sheetName).rowsToAdd.push(newRowData);
                        sheetInfo.dataMap.set(String(key), { rowIndex: -1, data: newRowData }); // Prevent duplicate adds
                    }
                }
            });

            // 6. Perform bulk write operations for each sheet
            for (const [sheetName, writes] of writesBySheet.entries()) {
                if (writes.rowsToUpdate.length === 0 && writes.rowsToAdd.length === 0) continue;
                sendMessageToDialog(`Writing ${writes.rowsToUpdate.length} updates and ${writes.rowsToAdd.length} new rows to "${sheetName}".`);

                const sheet = sheetCache.get(sheetName);
                const sheetInfo = sheetDataCache.get(sheetName);

                if (writes.rowsToUpdate.length > 0) {
                    for (const row of writes.rowsToUpdate) {
                        sheet.getRangeByIndexes(row.rowIndex, 0, 1, sheetInfo.headers.length).values = [row.values];
                    }
                }

                if (writes.rowsToAdd.length > 0) {
                    const startRow = sheetInfo.rowCount;
                    sheet.getRangeByIndexes(startRow, 0, writes.rowsToAdd.length, sheetInfo.headers.length).values = writes.rowsToAdd;
                    sheetInfo.rowCount += writes.rowsToAdd.length;
                }
                sheet.getUsedRange().format.autofitColumns();
            }
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

        if (!importFile.importName || typeof importFile.importName !== 'string') throw new Error("Missing a valid 'importName'.");
        if (!importFile.targetSheet || typeof importFile.targetSheet !== 'string') throw new Error("Missing a valid default 'targetSheet'.");
        if (!importFile.sheetKeyColumn || typeof importFile.sheetKeyColumn !== 'string') throw new Error("Missing a valid 'sheetKeyColumn'.");
        if (!Array.isArray(importFile.columnMappings) || importFile.columnMappings.length === 0) throw new Error("Must have a non-empty 'columnMappings' array.");

        for (const mapping of importFile.columnMappings) {
            if (!mapping.source || typeof mapping.source !== 'string' || !mapping.target || typeof mapping.target !== 'string') {
                throw new Error("Each item in 'columnMappings' must have valid 'source' and 'target' properties.");
            }
            if (mapping.targetSheet && typeof mapping.targetSheet !== 'string') {
                throw new Error("If 'targetSheet' is specified in a mapping, it must be a string.");
            }
        }
        
        if (!Array.isArray(importFile.data)) throw new Error("The 'data' property must be an array.");

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

