/*
 * This file contains the logic for the "Create LDA" ribbon button command.
 */
import { CONSTANTS, getSettings, findColumnIndex } from './utils.js';

let createLdaDialog = null;

/**
 * Checks for Master List and opens a dialog to guide the user.
 * @param {Office.AddinCommands.Event} event
 */
export async function openCreateLdaDialog(event) {
    let masterListExists = false;
    try {
        await Excel.run(async (context) => {
            // Check if "Master List" sheet exists
            const sheetNames = context.workbook.worksheets.load("items/name");
            await context.sync();
            for (let i = 0; i < sheetNames.items.length; i++) {
                if (sheetNames.items[i].name === CONSTANTS.MASTER_LIST_SHEET) {
                    masterListExists = true;
                    break;
                }
            }
        });
    } catch (error) {
        console.error("Error checking for Master List sheet:", error);
        // Assume it doesn't exist if there's an error
        masterListExists = false;
    }

    Office.context.ui.displayDialogAsync(
        'https://vsblanco.github.io/Student-Retention-Add-in/commands/create-lda-dialog.html',
        { height: 35, width: 35, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Create LDA dialog failed to open: " + asyncResult.error.message);
                event.completed();
                return;
            }
            createLdaDialog = asyncResult.value;
            createLdaDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                 const messageFromDialog = JSON.parse(arg.message);
                 if (messageFromDialog.type === 'dialogReady') {
                     // Dialog is ready, now send the result of our check.
                     if (createLdaDialog) {
                        createLdaDialog.messageChild(JSON.stringify({ type: 'masterListCheckResult', exists: masterListExists }));
                     }
                 } else {
                     // Handle other messages like 'createLdaSheet' or 'closeDialog'
                     processCreateLdaMessage(arg);
                 }
            });
            event.completed();
        }
    );
}


/**
 * Handles messages from the create LDA dialog.
 * @param {Office.DialogMessageReceivedEventArgs} arg
 */
async function processCreateLdaMessage(arg) {
    const message = JSON.parse(arg.message);
    switch (message.type) {
        case 'createLdaSheet':
            await handleCreateLdaSheet();
            break;
        case 'closeDialog':
            if (createLdaDialog) {
                createLdaDialog.close();
                createLdaDialog = null;
            }
            break;
        default:
            console.error("Unknown message type from create LDA dialog:", message.type);
            if (createLdaDialog) {
                createLdaDialog.close();
            }
    }
}

/**
 * Creates a new worksheet with today's date for LDA, populated with filtered and sorted data from the Master List.
 */
async function handleCreateLdaSheet() {
    console.log("[DEBUG] Starting handleCreateLdaSheet v13");
    try {
        const settings = await getSettings();
        const { daysOutFilter, includeFailingList, ldaColumns, hideLeftoverColumns } = settings.createlda;
        
        console.log(`[DEBUG] Settings: DaysOut=${daysOutFilter}, FailingList=${includeFailingList}, HideUnused=${hideLeftoverColumns}`);
        if (!ldaColumns || ldaColumns.length === 0) {
            console.warn("No columns selected for LDA report. Aborting.");
            if (createLdaDialog) createLdaDialog.messageChild(JSON.stringify({ type: 'creationError', error: "No columns selected in settings." }));
            return;
        }

        await Excel.run(async (context) => {
            console.log("[DEBUG] Reading data from Master List.");
            const masterSheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("values, formulas, rowIndex, columnIndex");
            await context.sync();

            const masterData = masterRange.values;
            const masterFormulas = masterRange.formulas;
            const originalHeaders = masterData[0];
            const originalLCHeaders = originalHeaders.map(h => String(h || '').toLowerCase());
            
            const daysOutColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);
            if (daysOutColIdx === -1) throw new Error("'Days Out' column not found in Master List.");

            // --- Start: New Color Caching Logic ---
            const valueToColorMap = new Map();
            const masterAssignedCol = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.assigned);
            if (masterAssignedCol !== -1) {
                const allAssignedValues = masterRange.values.map(row => row[masterAssignedCol]);
                const uniqueValues = [...new Set(allAssignedValues.slice(1).filter(v => v && String(v).trim() !== ""))];
                
                if (uniqueValues.length > 0) {
                    console.log(`[DEBUG] Found ${uniqueValues.length} unique values in 'Assigned' column. Fetching their colors...`);
                    const cellsToLoad = [];
                    uniqueValues.forEach(value => {
                        const firstInstanceIndex = allAssignedValues.indexOf(value);
                        if (firstInstanceIndex > 0) { // Ensure it's not the header
                            const absoluteRowIndex = masterRange.rowIndex + firstInstanceIndex;
                            const absoluteColIndex = masterRange.columnIndex + masterAssignedCol;
                            const cell = masterSheet.getCell(absoluteRowIndex, absoluteColIndex);
                            cell.load("format/fill/color");
                            cellsToLoad.push({ value: value, cell: cell });
                        }
                    });

                    await context.sync();

                    cellsToLoad.forEach(item => {
                        const color = item.cell.format.fill.color;
                        if (color && color !== '#ffffff' && color !== '#000000') {
                            valueToColorMap.set(item.value, color);
                        }
                    });
                    console.log(`[DEBUG] Cached colors for ${valueToColorMap.size} unique values.`);
                }
            }
            // --- End: New Color Caching Logic ---
            
            const dataRowsWithIndex = masterData.slice(1).map((row, index) => ({ row, originalIndex: index + 1 }));

            const filteredRows = dataRowsWithIndex.filter(({ row }) => {
                const daysOut = row[daysOutColIdx];
                return typeof daysOut === 'number' && daysOut >= daysOutFilter;
            });
            filteredRows.sort((a, b) => (b.row[daysOutColIdx] || 0) - (a.row[daysOutColIdx] || 0));
            
            const today = new Date();
            const baseSheetName = `LDA ${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
            let sheetName = baseSheetName;
            let counter = 2;
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();
            const existingSheetNames = new Set(worksheets.items.map(s => s.name));
            while (existingSheetNames.has(sheetName)) {
                sheetName = `${baseSheetName} (${counter++})`;
            }

            const newSheet = worksheets.add(sheetName);
            newSheet.activate();
            
            let ldaTableEndRow = await createAndFormatTable(context, {
                sheet: newSheet,
                tableName: sheetName.replace(/[^a-zA-Z0-9]/g, "_") + "_LDA",
                startRow: 0,
                dataRows: filteredRows,
                masterFormulas,
                ldaColumns,
                hideLeftoverColumns,
                originalHeaders,
                valueToColorMap // Pass the color map
            });
            
            if (includeFailingList) {
                console.log("[DEBUG] includeFailingList is true, creating failing list.");
                const nextStartRow = ldaTableEndRow > 0 ? ldaTableEndRow + 2 : 3;
                await createFailingListTable(context, newSheet, sheetName, nextStartRow, dataRowsWithIndex, masterFormulas, ldaColumns, hideLeftoverColumns, originalHeaders, daysOutColIdx, valueToColorMap);
            }
            
        });

        if (createLdaDialog) {
            createLdaDialog.messageChild(JSON.stringify({ type: 'creationSuccess' }));
        }
        console.log("[DEBUG] handleCreateLdaSheet completed successfully.");

    } catch (error) {
        console.error("Error creating LDA sheet: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
        if (createLdaDialog) {
            createLdaDialog.messageChild(JSON.stringify({ type: 'creationError', error: error.message }));
        }
    }
}

async function createFailingListTable(context, sheet, sheetName, startRow, masterDataWithIndex, masterFormulas, ldaColumns, hideLeftoverColumns, originalHeaders, daysOutColIdx, valueToColorMap) {
    console.log("[DEBUG] Creating failing list table.");
    const originalLCHeaders = originalHeaders.map(h => String(h || '').toLowerCase());
    const gradeColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);

    if (gradeColIdx === -1) {
        console.warn("'Grade' column not found, cannot create failing list.");
        return;
    }
    
    if (daysOutColIdx === -1) {
        console.warn("'Days Out' column not found, cannot filter failing list by days out.");
        return;
    }

    const failingRows = masterDataWithIndex.filter(({ row }) => {
        const grade = row[gradeColIdx];
        const daysOut = row[daysOutColIdx];
        const isFailingGrade = typeof grade === 'number' && (grade < 0.60 || (grade >= 1 && grade < 60));
        const isRecent = typeof daysOut === 'number' && daysOut <= 4;
        return isFailingGrade && isRecent;
    });
    failingRows.sort((a, b) => (a.row[gradeColIdx] || 0) - (b.row[gradeColIdx] || 0));

    if (failingRows.length > 0) {
        const titleRange = sheet.getRangeByIndexes(startRow, 0, 1, 1);
        titleRange.values = [["Failing Students (Active)"]];
        titleRange.format.font.bold = true;
        
        await createAndFormatTable(context, {
            sheet,
            tableName: sheetName.replace(/[^a-zA-Z0-9]/g, "_") + "_Failing",
            startRow: startRow + 1,
            dataRows: failingRows,
            masterFormulas,
            ldaColumns,
            hideLeftoverColumns,
            originalHeaders,
            valueToColorMap // Pass the color map
        });
    }
}
async function applyGradeConditionalFormattingToTable(context, table) {
    // 1. Load table properties needed
    table.load("name, showHeaders");
    await context.sync();

    if (!table.showHeaders) {
        console.log(`Table '${table.name}' has no headers, skipping conditional formatting.`);
        return;
    }
    
    // 2. Get headers from the table
    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");
    await context.sync();
    
    const lowerCaseHeaders = headerRange.values[0].map(h => String(h || '').toLowerCase());
    const gradeColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);

    if (gradeColIdx === -1) {
        console.log(`'Grade' column not found in table '${table.name}', skipping conditional formatting.`);
        return;
    }

    // 3. Get the data body range for the grade column
    const gradeColumn = table.columns.getItemAt(gradeColIdx);
    const gradeColumnRange = gradeColumn.getDataBodyRange();
    gradeColumnRange.load("values, rowCount");
    await context.sync();

    if (gradeColumnRange.rowCount === 0) {
        console.log(`Table '${table.name}' has no data rows to format.`);
        return;
    }

    // 4. Determine scale (0-1 or 0-100)
    let isPercentScale = false;
    for (let i = 0; i < Math.min(gradeColumnRange.rowCount, 10); i++) {
        if (gradeColumnRange.values[i] && typeof gradeColumnRange.values[i][0] === 'number' && gradeColumnRange.values[i][0] > 1) {
            isPercentScale = true;
            break;
        }
    }
    
    console.log(`Detected grade scale for table '${table.name}': ${isPercentScale ? '0-100' : '0-1'}. Applying 3-color scale.`);

    // 5. Apply formatting
    gradeColumnRange.conditionalFormats.clearAll();
    const conditionalFormat = gradeColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    const criteria = {
        minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" }, // Red
        midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: isPercentScale ? "70" : "0.7", color: "#FFEB84" }, // Yellow
        maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" } // Green
    };
    conditionalFormat.colorScale.criteria = criteria;

    await context.sync();
    console.log(`Conditional formatting applied to table '${table.name}'.`);
}


async function createAndFormatTable(context, options) {
    const {
        sheet,
        tableName,
        startRow,
        dataRows,
        masterFormulas,
        ldaColumns,
        hideLeftoverColumns,
        originalHeaders,
        valueToColorMap
    } = options;

    let finalHeaders;
    if (hideLeftoverColumns) {
        // Start with the user-defined order
        finalHeaders = [...ldaColumns];
        const ldaColumnsSet = new Set(ldaColumns);
        // Add any other columns from the master list that weren't in the user's list
        originalHeaders.forEach(h => {
            if (!ldaColumnsSet.has(h)) {
                finalHeaders.push(h);
            }
        });
    } else {
        // Only use the columns specified in settings
        finalHeaders = ldaColumns;
    }

    const indicesToKeep = finalHeaders.map(h => originalHeaders.indexOf(h));
    const originalLCHeaders = originalHeaders.map(h => String(h ?? '').toLowerCase());
    const gradeBookColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);

    const dataToWrite = [];
    const formulasToWrite = [];

    dataRows.forEach(({
        row,
        originalIndex
    }) => {
        const newRow = [];
        const formulaRow = new Array(finalHeaders.length).fill(null);

        indicesToKeep.forEach((keptIdx, i) => {
            newRow.push(row[keptIdx] ?? "");
            if (keptIdx === gradeBookColIdx) {
                const formula = masterFormulas[originalIndex][keptIdx];
                const value = row[keptIdx];
                if (typeof formula === 'string' && formula.toLowerCase().startsWith('=hyperlink')) {
                    formulaRow[i] = formula;
                    const match = formula.match(/, *"([^"]+)"\)/i);
                    newRow[i] = match ? match[1] : "Gradebook";
                } else if (typeof value === 'string' && (value.startsWith('http://') || value.startsWith('https://'))) {
                    formulaRow[i] = `=HYPERLINK("${value}", "Gradebook")`;
                    newRow[i] = "Gradebook";
                }
            }
        });

        dataToWrite.push(newRow);
        formulasToWrite.push(formulaRow);
    });

    const finalValues = [finalHeaders, ...dataToWrite];
    const finalFormulas = [new Array(finalHeaders.length).fill(null), ...formulasToWrite];

    if (finalValues.length > 1) {
        const dataRange = sheet.getRangeByIndexes(startRow, 0, finalValues.length, finalHeaders.length);
        dataRange.values = finalValues;
        dataRange.formulas = finalFormulas;

        const table = sheet.tables.add(dataRange, true);
        table.name = tableName;
        table.style = "TableStyleLight9";
        
        await context.sync(); // Sync after table creation

        // --- Start: Apply Preserved Colors ---
        if (valueToColorMap.size > 0) {
            const assignedColIdxInTable = findColumnIndex(finalHeaders.map(h => h.toLowerCase()), CONSTANTS.COLUMN_MAPPINGS.assigned);
            if (assignedColIdxInTable !== -1) {
                const tableBodyRange = table.getDataBodyRange();
                tableBodyRange.load("values");
                await context.sync();
                
                for (let i = 0; i < tableBodyRange.values.length; i++) {
                    const assignedValue = tableBodyRange.values[i][assignedColIdxInTable];
                    if (assignedValue && valueToColorMap.has(assignedValue)) {
                        const color = valueToColorMap.get(assignedValue);
                        const cellToColor = table.getDataBodyRange().getCell(i, assignedColIdxInTable);
                        cellToColor.format.fill.color = color;
                    }
                }
            }
        }
        // --- End: Apply Preserved Colors ---

        await applyGradeConditionalFormattingToTable(context, table);
      
        sheet.getUsedRange().getEntireColumn().format.autofitColumns();
        await context.sync();
      
        if (hideLeftoverColumns) {
            console.log("[DEBUG] Hiding unused columns for table:", tableName);
            const selectedColumnsSet = new Set(ldaColumns.map(h => h.toLowerCase()));

            table.columns.load("items/name");
            await context.sync();

            table.columns.items.forEach((col, idx) => {
                const colName = col.name.trim().toLowerCase();
                if (!selectedColumnsSet.has(colName)) {
                    try {
                        console.log(`[DEBUG] Hiding worksheet column at index ${idx} for "${col.name}"`);
                        const columnLetter = String.fromCharCode(65 + idx); // A, B, C, ...
                        const worksheetColumn = sheet.getRange(`${columnLetter}:${columnLetter}`);
                        console.log(`[DEBUG] Column index ${idx} maps to letter ${columnLetter} for column "${col.name}"`);
                        worksheetColumn.columnHidden = true;

                    } catch (err) {
                        console.warn(`Failed to hide worksheet column for "${col.name}":`, err);
                    }
                }
            });
            await context.sync();
        }


        const dateColumnsToFormat = ["lda", "dod", "expstartdate"];
        dateColumnsToFormat.forEach(colName => {
            const colIdx = findColumnIndex(finalHeaders.map(h => h.toLowerCase()), [colName.toLowerCase()]);
            if (colIdx !== -1) {
                table.columns.getItemAt(colIdx).getRange().numberFormat = [
                    ["m/d/yyyy"]
                ];
            }
        });

        return startRow + finalValues.length;
    } else {
        sheet.getRangeByIndexes(startRow, 0, 1, finalHeaders.length).values = [finalHeaders];
        return startRow + 1;
    }
}
