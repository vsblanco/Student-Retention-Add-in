// Timestamp: 2025-09-23 11:12 AM EDT
// Version: 2.6.0
/*
 * This file contains the logic for handling custom, schema-driven imports
 * from a single JSON file.
 * Version 2.6.0 adds an 'overwriteTargetSheet' option to the schema.
 * When true, it removes rows from the primary target sheet that do not
 * exist in the source JSON file, ensuring an exact match.
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
 * Extracts the key from a cell, prioritizing the URL from a HYPERLINK formula.
 * @param {string} formula The formula of the cell.
 * @param {any} value The value of the cell.
 * @returns {string|null} The extracted key as a string, or null.
 */
function getKeyFromCell(formula, value) {
    if (typeof formula === 'string' && formula.toUpperCase().startsWith('=HYPERLINK')) {
        const match = formula.match(/=HYPERLINK\("([^"]+)"/i);
        if (match && match[1]) {
            return match[1]; // Return the URL part of the hyperlink
        }
    }
    return value ? String(value) : null; // Fallback to the cell's display value
}


/**
 * Main function to handle the custom import process from a self-contained JSON file.
 * @param {object} message The message from the dialog, containing the file data.
 * @param {function} sendMessageToDialog A function to send status messages back to the dialog.
 */
export async function handleCustomImport(message, sendMessageToDialog) {
    try {
        sendMessageToDialog("Starting custom import from JSON file...");

        const importPayload = parseAndValidateImportFile(message.data, sendMessageToDialog);
        if (!importPayload) return;

        const { schema, data: sourceData } = importPayload;
        sendMessageToDialog(`Successfully validated schema: "${schema.importName}".`);
        if (schema.overwriteTargetSheet) {
            sendMessageToDialog(`Overwrite mode enabled for sheet: "${schema.targetSheet}".`);
        }
        sendMessageToDialog(`Found ${sourceData.length} flattened rows to process.`);

        const allSheetNames = new Set([schema.targetSheet]);
        schema.columnMappings.forEach(m => { if (m.targetSheet) allSheetNames.add(m.targetSheet); });
        sendMessageToDialog(`Preparing to process ${allSheetNames.size} target sheet(s).`);

        await Excel.run(async (context) => {
            const sheetCache = new Map();
            const sheetDataCache = new Map();

            // 1. Ensure all sheets exist and cache them
            for (const sheetName of allSheetNames) {
                let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                await context.sync();
                if (sheet.isNullObject) {
                    sheet = context.workbook.worksheets.add(sheetName);
                    const primaryKey = Array.isArray(schema.sheetKeyColumn) ? schema.sheetKeyColumn[0] : schema.sheetKeyColumn;
                    
                    const headerArray = [primaryKey];
                    const seenHeaders = new Set([primaryKey]);
                    schema.columnMappings
                        .filter(m => (m.targetSheet || schema.targetSheet) === sheetName)
                        .forEach(m => {
                            const targetHeader = Array.isArray(m.target) ? m.target[0] : m.target;
                            if (!seenHeaders.has(targetHeader)) {
                                headerArray.push(targetHeader);
                                seenHeaders.add(targetHeader);
                            }
                        });
                    sheet.getRangeByIndexes(0, 0, 1, headerArray.length).values = [headerArray];
                }
                sheetCache.set(sheetName, sheet);
            }
            await context.sync();

            // 2. Read all sheet data into memory and create maps
            for (const [sheetName, sheet] of sheetCache.entries()) {
                const usedRange = sheet.getUsedRange(true);
                usedRange.load("values, formulas, rowCount");
                await context.sync();

                const formulas = usedRange.formulas || [];
                const headers = formulas.length > 0 ? formulas[0].map(h => String(h || '')) : [];
                const resolvedKeyColumn = findFirstMatchingHeader(schema.sheetKeyColumn, headers);
                if (!resolvedKeyColumn) throw new Error(`Key column not found in sheet "${sheetName}".`);
                const keyColumnIndex = headers.indexOf(resolvedKeyColumn);
                
                const dataMap = new Map();
                for (let i = 1; i < usedRange.rowCount; i++) {
                    const key = getKeyFromCell(formulas[i][keyColumnIndex], formulas[i][keyColumnIndex]);
                    if (key) dataMap.set(String(key), { rowIndex: i, formulas: formulas[i] });
                }
                
                sheetDataCache.set(sheetName, { headers, resolvedKeyColumn, keyColumnIndex, dataMap, formulas: formulas, rowCount: usedRange.rowCount });
                sendMessageToDialog(`Mapped ${dataMap.size} existing rows for sheet "${sheetName}".`);
            }

            // 3. Resolve source keys for each sheet
            const sourceKeyMappingBySheet = new Map();
            for (const [sheetName, sheetInfo] of sheetDataCache.entries()) {
                const keyColumnName = sheetInfo.resolvedKeyColumn;
                
                const mapping = schema.columnMappings.find(m =>
                    (m.targetSheet || schema.targetSheet) === sheetName &&
                    (Array.isArray(m.target) ? m.target : [m.target]).includes(keyColumnName)
                );

                if (!mapping) throw new Error(`Could not find a 'source' in columnMappings for the key column "${keyColumnName}" on sheet "${sheetName}".`);
                sourceKeyMappingBySheet.set(sheetName, mapping.source);
            }

            // Prepare a set of all source keys for the primary sheet if overwrite is enabled
            let mainSheetSourceKeySet = null;
            if (schema.overwriteTargetSheet) {
                const mainSheetSourceKey = sourceKeyMappingBySheet.get(schema.targetSheet);
                mainSheetSourceKeySet = new Set(sourceData.map(item => String(item[mainSheetSourceKey])));
            }


            // 4. Prepare updates and new rows IN MEMORY
            const writesBySheet = new Map();
            for (const sheetName of allSheetNames) {
                const sheetInfo = sheetDataCache.get(sheetName);
                writesBySheet.set(sheetName, {
                    finalFormulas: sheetInfo.formulas.map(row => [...row]),
                    rowsToAdd: [],
                    updatedRowCount: 0
                });
            }

            sourceData.forEach(sourceObject => {
                const valuesBySheet = new Map();
                 schema.columnMappings.forEach(mapping => {
                    const sheetName = mapping.targetSheet || schema.targetSheet;
                    if (!valuesBySheet.has(sheetName)) valuesBySheet.set(sheetName, {});
                    const sheetInfo = sheetDataCache.get(sheetName);
                    if (!sheetInfo) return; 
                    const targetHeader = findFirstMatchingHeader(mapping.target, sheetInfo.headers);
                    if (targetHeader && sourceObject[mapping.source] !== undefined) {
                        valuesBySheet.get(sheetName)[targetHeader] = sourceObject[mapping.source];
                    }
                });

                for (const [sheetName, newValues] of valuesBySheet.entries()) {
                    const sheetInfo = sheetDataCache.get(sheetName);
                    const writes = writesBySheet.get(sheetName);
                    const sourceKeyName = sourceKeyMappingBySheet.get(sheetName);
                    const key = sourceObject[sourceKeyName];
                    if (key === undefined) continue;

                    const existingRow = sheetInfo.dataMap.get(String(key));
                    if (existingRow) {
                        const rowIndex = existingRow.rowIndex;
                        const targetRow = writes.finalFormulas[rowIndex];

                        if (!targetRow) {
                            console.warn(`Inconsistent row index found for key: ${key}`);
                            continue;
                        }

                        for (let i = 0; i < sheetInfo.headers.length; i++) {
                            const header = sheetInfo.headers[i];
                            const isKeyCol = (i === sheetInfo.keyColumnIndex);
                            const isFormula = typeof targetRow[i] === 'string' && targetRow[i].startsWith('=');
                            if (isKeyCol && isFormula) continue;
                            if (newValues[header] !== undefined) {
                                targetRow[i] = newValues[header];
                            }
                        }
                         writes.updatedRowCount++;
                    } else {
                        const newRow = new Array(sheetInfo.headers.length).fill(null);
                        sheetInfo.headers.forEach((header, index) => {
                            if (newValues[header] !== undefined) newRow[index] = newValues[header];
                        });
                        newRow[sheetInfo.keyColumnIndex] = key;
                        writes.rowsToAdd.push(newRow);
                        sheetInfo.dataMap.set(String(key), { rowIndex: -1 });
                    }
                }
            });

            // 5. Perform BULK write operations
            for (const [sheetName, writes] of writesBySheet.entries()) {
                // --- Overwrite logic for the primary target sheet ---
                if (sheetName === schema.targetSheet && schema.overwriteTargetSheet) {
                    const sheet = sheetCache.get(sheetName);
                    const sheetInfo = sheetDataCache.get(sheetName);
                    const keyIndex = sheetInfo.keyColumnIndex;

                    // Filter the existing (and updated) rows to keep only those present in the source JSON
                    const rowsToKeep = [writes.finalFormulas[0]]; // Always keep the header row
                    for (let i = 1; i < writes.finalFormulas.length; i++) {
                        const row = writes.finalFormulas[i];
                        const key = getKeyFromCell(row[keyIndex], row[keyIndex]);
                        if (mainSheetSourceKeySet.has(String(key))) {
                            rowsToKeep.push(row);
                        }
                    }

                    const finalSheetContent = [...rowsToKeep, ...writes.rowsToAdd];
                    const rowsRemovedCount = (sheetInfo.rowCount - 1) - (rowsToKeep.length - 1);
                    
                    sendMessageToDialog(`For sheet "${sheetName}": ${rowsToKeep.length -1} rows kept, ${writes.rowsToAdd.length} rows added, ${rowsRemovedCount} rows removed.`);

                    // Clear the sheet and write the final, synchronized content
                    const usedRange = sheet.getUsedRange();
                    if(usedRange) usedRange.clear();
                    await context.sync();
                    
                    if (finalSheetContent.length > 0) {
                        const newRange = sheet.getRangeByIndexes(0, 0, finalSheetContent.length, sheetInfo.headers.length);
                        newRange.values = finalSheetContent;
                    }
                    sheet.getUsedRange().format.autofitColumns();
                    continue; // Skip the standard write logic for this sheet
                }
                
                // --- Standard update/add logic for all other sheets ---
                if (writes.updatedRowCount === 0 && writes.rowsToAdd.length === 0) continue;
                sendMessageToDialog(`Writing ${writes.updatedRowCount} updates and ${writes.rowsToAdd.length} new rows to "${sheetName}".`);
                
                const sheet = sheetCache.get(sheetName);
                const sheetInfo = sheetDataCache.get(sheetName);

                if (writes.updatedRowCount > 0) {
                    const updateRange = sheet.getRangeByIndexes(0, 0, sheetInfo.rowCount, sheetInfo.headers.length);
                    updateRange.values = writes.finalFormulas;
                }
                
                if (writes.rowsToAdd.length > 0) {
                    const addRange = sheet.getRangeByIndexes(sheetInfo.rowCount, 0, writes.rowsToAdd.length, sheetInfo.headers.length);
                    addRange.values = writes.rowsToAdd;
                }
                sheet.getUsedRange().format.autofitColumns();
            }
            await context.sync();
        });

        sendMessageToDialog("Custom import completed successfully.", 'complete');

    } catch (error) {
        console.error("Custom Import Error: ", error);
        sendMessageToDialog(`Error: ${error.message}`, 'error', [error.stack]);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug Info: " + JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Parses and validates the provided JSON import file.
 * It now handles encapsulated schemas and flattens nested data arrays.
 * @param {string} dataUrl The data URL of the JSON file.
 * @param {function} sendMessageToDialog Function to send status messages.
 * @returns {object|null} An object with 'schema' and 'data' properties, or null if validation fails.
 */
function parseAndValidateImportFile(dataUrl, sendMessageToDialog) {
    try {
        const jsonString = atob(dataUrl.split(',')[1]);
        let importFile = JSON.parse(jsonString);

        if (importFile.CUSTOM_IMPORT && typeof importFile.CUSTOM_IMPORT === 'object') {
            sendMessageToDialog("Detected encapsulated 'CUSTOM_IMPORT' schema.");
            importFile = importFile.CUSTOM_IMPORT;
        }

        if (!importFile.importName || typeof importFile.importName !== 'string') throw new Error("Missing a valid 'importName'.");
        if (!importFile.targetSheet || typeof importFile.targetSheet !== 'string') throw new Error("Missing a valid default 'targetSheet'.");
        if (!importFile.sheetKeyColumn || (typeof importFile.sheetKeyColumn !== 'string' && !Array.isArray(importFile.sheetKeyColumn))) throw new Error("'sheetKeyColumn' must be a non-empty string or an array of non-empty strings.");
        if (Array.isArray(importFile.sheetKeyColumn) && importFile.sheetKeyColumn.length === 0) throw new Error("'sheetKeyColumn' array cannot be empty.");
        if (!Array.isArray(importFile.columnMappings) || importFile.columnMappings.length === 0) throw new Error("Must have a non-empty 'columnMappings' array.");
        if (importFile.overwriteTargetSheet && typeof importFile.overwriteTargetSheet !== 'boolean') throw new Error("If 'overwriteTargetSheet' is specified, it must be a boolean (true/false).");

        for (const mapping of importFile.columnMappings) {
            if (!mapping.source || typeof mapping.source !== 'string') throw new Error("Each item in 'columnMappings' must have a valid 'source' property.");
            if (!mapping.target || (typeof mapping.target !== 'string' && !Array.isArray(mapping.target))) throw new Error("Each 'target' in 'columnMappings' must be a non-empty string or an array of non-empty strings.");
            if (Array.isArray(mapping.target) && mapping.target.length === 0) throw new Error("A 'target' array cannot be empty.");
            if (mapping.targetSheet && typeof mapping.targetSheet !== 'string') throw new Error("If 'targetSheet' is specified in a mapping, it must be a string.");
        }
        if (!Array.isArray(importFile.data)) throw new Error("The 'data' property must be an array.");

        const schema = {
            importName: importFile.importName,
            targetSheet: importFile.targetSheet,
            sheetKeyColumn: importFile.sheetKeyColumn,
            columnMappings: importFile.columnMappings,
            dataArrayKey: importFile.dataArrayKey,
            overwriteTargetSheet: importFile.overwriteTargetSheet || false
        };

        let sourceData = importFile.data;
        if (schema.dataArrayKey && typeof schema.dataArrayKey === 'string') {
            sendMessageToDialog(`Flattening nested data using key: '${schema.dataArrayKey}'...`);
            const flattenedData = [];
            sourceData.forEach(parentObject => {
                const nestedArray = parentObject[schema.dataArrayKey];
                
                if (Array.isArray(nestedArray) && nestedArray.length > 0) {
                    nestedArray.forEach(childObject => {
                        const combinedObject = { ...parentObject, ...childObject };
                        delete combinedObject[schema.dataArrayKey];
                        flattenedData.push(combinedObject);
                    });
                } else {
                    const parentOnlyObject = { ...parentObject };
                    delete parentOnlyObject[schema.dataArrayKey]; 
                    flattenedData.push(parentOnlyObject);
                }
            });
            sourceData = flattenedData;
        }

        return { schema, data: sourceData };

    } catch (error) {
        sendMessageToDialog(`JSON file validation failed: ${error.message}`, 'error');
        return null;
    }
}
