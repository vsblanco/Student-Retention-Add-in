// Timestamp: 2025-09-18 01:43 PM EDT
// Version: 1.5.0
/*
 * This file contains the logic for handling custom, schema-driven imports
 * from a single JSON file that contains both the schema and the data.
 * Version 1.5.0 introduces support for fallback names for the 'sheetKeyColumn'
 * by allowing it to be an array of strings.
 */

/**
 * Finds the first valid header from a target (string or array) that exists in the sheet headers.
 * @param {string|string[]} target The target property from the schema mapping.
 * @param {string[]} sheetHeaders The array of headers from the actual Excel sheet.
 * @returns {string|null} The matched header name or null if no match is found.
 */
function findFirstMatchingHeader(target, sheetHeaders) {
    const targets = Array.isArray(target) ? target : [target];
    for (const t of targets) {
        if (sheetHeaders.includes(t)) {
            return t;
        }
    }
    return null;
}

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

        const allSheetNames = new Set([schema.targetSheet]);
        schema.columnMappings.forEach(m => {
            if (m.targetSheet) allSheetNames.add(m.targetSheet);
        });
        sendMessageToDialog(`Preparing to process ${allSheetNames.size} target sheet(s).`);

        await Excel.run(async (context) => {
            const sheetCache = new Map();
            const sheetDataCache = new Map();

            const primaryKeyColumn = Array.isArray(schema.sheetKeyColumn) ? schema.sheetKeyColumn[0] : schema.sheetKeyColumn;

            // 3. Ensure all target sheets exist
            for (const sheetName of allSheetNames) {
                let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                await context.sync();

                if (sheet.isNullObject) {
                    sendMessageToDialog(`Sheet "${sheetName}" not found, creating it...`);
                    sheet = context.workbook.worksheets.add(sheetName);
                    
                    const headers = new Set([primaryKeyColumn]);
                    schema.columnMappings
                        .filter(m => (m.targetSheet || schema.targetSheet) === sheetName)
                        .forEach(m => {
                            const primaryTarget = Array.isArray(m.target) ? m.target[0] : m.target;
                            headers.add(primaryTarget);
                        });
                    
                    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.size);
                    headerRange.values = [Array.from(headers)];
                    headerRange.format.font.bold = true;
                } else {
                    sendMessageToDialog(`Found existing sheet: "${sheetName}".`);
                }
                sheetCache.set(sheetName, sheet);
            }
            await context.sync();

            // 4. Read existing data and resolve the key column for each sheet
            for (const [sheetName, sheet] of sheetCache.entries()) {
                const usedRange = sheet.getUsedRange(true);
                usedRange.load("values, rowCount");
                await context.sync();

                const values = usedRange.values || [];
                const headers = values.length > 0 ? values[0].map(h => String(h || '')) : [];
                
                const resolvedKeyColumn = findFirstMatchingHeader(schema.sheetKeyColumn, headers);
                if (!resolvedKeyColumn) {
                    throw new Error(`None of the specified key columns were found in sheet "${sheetName}". Please add one and try again.`);
                }
                const keyColumnIndex = headers.indexOf(resolvedKeyColumn);

                const dataMap = new Map();
                for (let i = 1; i < usedRange.rowCount; i++) {
                    const key = values[i][keyColumnIndex];
                    if (key) dataMap.set(String(key), { rowIndex: i, data: values[i] });
                }
                sheetDataCache.set(sheetName, { headers, resolvedKeyColumn, keyColumnIndex, dataMap, rowCount: usedRange.rowCount });
                sendMessageToDialog(`Using "${resolvedKeyColumn}" as key for sheet "${sheetName}". Mapped ${dataMap.size} existing rows.`);
            }

            // 5. Find the single source key from the data file
            const keyColumnNames = Array.isArray(schema.sheetKeyColumn) ? schema.sheetKeyColumn : [schema.sheetKeyColumn];
            let sourceKey = null;
            const sourceKeyMapping = schema.columnMappings.find(m => {
                const targets = Array.isArray(m.target) ? m.target : [m.target];
                return targets.some(t => keyColumnNames.includes(t));
            });
            if (!sourceKeyMapping) throw new Error(`Could not find a 'source' in columnMappings for any of the specified sheetKeyColumns.`);
            sourceKey = sourceKeyMapping.source;

            // 6. Prepare data for writing
            const writesBySheet = new Map();
            allSheetNames.forEach(name => writesBySheet.set(name, { rowsToUpdate: [], rowsToAdd: [] }));

            sourceData.forEach(sourceObject => {
                const key = sourceObject[sourceKey];
                if (key === undefined) return;

                const valuesBySheet = new Map();
                schema.columnMappings.forEach(mapping => {
                    const sheetName = mapping.targetSheet || schema.targetSheet;
                    const sheetInfo = sheetDataCache.get(sheetName);
                    if (!sheetInfo) return;

                    if (!valuesBySheet.has(sheetName)) valuesBySheet.set(sheetName, {});
                    
                    const resolvedTargetHeader = findFirstMatchingHeader(mapping.target, sheetInfo.headers);

                    if (resolvedTargetHeader && sourceObject[mapping.source] !== undefined) {
                        valuesBySheet.get(sheetName)[resolvedTargetHeader] = sourceObject[mapping.source];
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
                        sheetInfo.dataMap.set(String(key), { rowIndex: -1, data: newRowData }); 
                    }
                }
            });

            // 7. Perform bulk write operations
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
        
        if (!importFile.sheetKeyColumn || (typeof importFile.sheetKeyColumn !== 'string' && !Array.isArray(importFile.sheetKeyColumn))) {
            throw new Error("'sheetKeyColumn' must be a non-empty string or an array of non-empty strings.");
        }
        if (Array.isArray(importFile.sheetKeyColumn) && importFile.sheetKeyColumn.length === 0) {
            throw new Error("'sheetKeyColumn' array cannot be empty.");
        }


        if (!Array.isArray(importFile.columnMappings) || importFile.columnMappings.length === 0) throw new Error("Must have a non-empty 'columnMappings' array.");

        for (const mapping of importFile.columnMappings) {
            if (!mapping.source || typeof mapping.source !== 'string') {
                 throw new Error("Each item in 'columnMappings' must have a valid 'source' property.");
            }
            if (!mapping.target || (typeof mapping.target !== 'string' && !Array.isArray(mapping.target))) {
                throw new Error("Each 'target' in 'columnMappings' must be a non-empty string or an array of non-empty strings.");
            }
            if (Array.isArray(mapping.target) && mapping.target.length === 0) {
                throw new Error("A 'target' array cannot be empty.");
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

