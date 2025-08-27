/*
 * This file contains the logic for the "Create LDA" ribbon button command.
 */
import { CONSTANTS, getSettings, findColumnIndex, parseDate } from './utils.js';

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
    console.log("[DEBUG] Starting handleCreateLdaSheet v17");
    try {
        const settings = await getSettings();
        const { daysOutFilter, includeFailingList, ldaColumns, hideLeftoverColumns, includeLdaTagFollowup } = settings.createlda;
        
        console.log(`[DEBUG] Settings: DaysOut=${daysOutFilter}, FailingList=${includeFailingList}, HideUnused=${hideLeftoverColumns}, LdaFollowup=${includeLdaTagFollowup}`);
        if (!ldaColumns || ldaColumns.length === 0) {
            console.warn("No columns selected for LDA report. Aborting.");
            if (createLdaDialog) createLdaDialog.messageChild(JSON.stringify({ type: 'creationError', error: "No columns selected in settings." }));
            return;
        }

        await Excel.run(async (context) => {
            // --- Start: Read Master List and History Data ---
            console.log("[DEBUG] Reading data from Master List.");
            const masterSheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("values, formulas, rowIndex, columnIndex");
            
            let historySheet, historyRange, historyData;
            const ldaFollowUpMap = new Map();

            // Always try to get history sheet for DNC check
            try {
                historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
                historyRange = historySheet.getUsedRange();
                historyRange.load("values");
            } catch (e) {
                console.warn("Student History sheet not found, cannot process LDA follow-ups or DNC statuses.");
            }
            
            await context.sync();
            
            // --- Process History for LDA Follow-ups and DNC ---
            const dncStatusMap = new Map();
            if (historySheet) {
                historyData = historyRange.values;
                console.log("[DEBUG] Processing Student History sheet.");
                const historyHeaders = historyData[0].map(h => String(h || '').toLowerCase());
                const histIdCol = findColumnIndex(historyHeaders, CONSTANTS.STUDENT_NUMBER_COLS);
                const histTagCol = findColumnIndex(historyHeaders, ["tag"]);

                if (histIdCol !== -1 && histTagCol !== -1) {
                    const today = new Date();
                    today.setHours(0, 0, 0, 0);

                    // Iterate backwards to get the most recent tags for each student
                    for (let i = historyData.length - 1; i > 0; i--) { 
                        const row = historyData[i];
                        const studentId = String(row[histIdCol]);
                        if (!studentId) continue;

                        // DNC Check (always run)
                        if (!dncStatusMap.has(studentId)) {
                             const tags = String(row[histTagCol] || '').toLowerCase().split(',').map(t => t.trim());
                             const dncTag = tags.find(t => t.startsWith('dnc'));
                             if (dncTag) {
                                 const dncInfo = { phone: false, email: false, all: false };
                                 if (dncTag.includes('phone')) dncInfo.phone = true;
                                 else if (dncTag.includes('email')) dncInfo.email = true;
                                 else dncInfo.all = true; // General "dnc" tag
                                 dncStatusMap.set(studentId, dncInfo);
                             }
                        }

                        // LDA Follow-up Check (if enabled)
                        if (includeLdaTagFollowup && !ldaFollowUpMap.has(studentId)) {
                            const tags = String(row[histTagCol] || '').split(',').map(t => t.trim());
                            const ldaTag = tags.find(t => t.toLowerCase().startsWith('lda '));
                            if (ldaTag) {
                                const dateStr = ldaTag.substring(4);
                                const ldaDate = new Date(dateStr);
                                if (!isNaN(ldaDate.getTime())) {
                                    ldaDate.setHours(0, 0, 0, 0);
                                    if (ldaDate >= today) {
                                        ldaFollowUpMap.set(studentId, { tag: ldaTag, date: ldaDate });
                                    }
                                }
                            }
                        }
                    }
                    console.log(`[DEBUG] Found ${ldaFollowUpMap.size} students with future LDA follow-ups.`);
                    console.log(`[DEBUG] Found ${dncStatusMap.size} students with DNC preferences.`);
                }
            }


            const masterData = masterRange.values;
            const masterFormulas = masterRange.formulas;
            const originalHeaders = masterData[0];
            const originalLCHeaders = originalHeaders.map(h => String(h || '').toLowerCase());
            
            const daysOutColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);
            if (daysOutColIdx === -1) throw new Error("'Days Out' column not found in Master List.");

            // --- Start: Color Caching Logic ---
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
            // --- End: Color Caching Logic ---
            
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
                valueToColorMap,
                ldaFollowUpMap,
                dncStatusMap // Pass DNC map
            });
            
            if (includeFailingList) {
                console.log("[DEBUG] includeFailingList is true, creating failing list.");
                const nextStartRow = ldaTableEndRow > 0 ? ldaTableEndRow + 2 : 3;
                await createFailingListTable(context, newSheet, sheetName, nextStartRow, dataRowsWithIndex, masterFormulas, ldaColumns, hideLeftoverColumns, originalHeaders, daysOutColIdx, valueToColorMap, ldaFollowUpMap, dncStatusMap);
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

async function createFailingListTable(context, sheet, sheetName, startRow, masterDataWithIndex, masterFormulas, ldaColumns, hideLeftoverColumns, originalHeaders, daysOutColIdx, valueToColorMap, ldaFollowUpMap, dncStatusMap) {
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
            valueToColorMap,
            ldaFollowUpMap,
            dncStatusMap // Pass DNC map
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
        valueToColorMap,
        ldaFollowUpMap,
        dncStatusMap
    } = options;

    let finalHeaders;
    if (hideLeftoverColumns) {
        finalHeaders = [...ldaColumns];
        const ldaColumnsSet = new Set(ldaColumns);
        originalHeaders.forEach(h => {
            if (!ldaColumnsSet.has(h)) {
                finalHeaders.push(h);
            }
        });
    } else {
        finalHeaders = ldaColumns;
    }

    const indicesToKeep = finalHeaders.map(h => originalHeaders.indexOf(h));
    const originalLCHeaders = originalHeaders.map(h => String(h ?? '').toLowerCase());
    const gradeBookColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
    const studentNumberColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.STUDENT_NUMBER_COLS);

    const dataToWrite = [];
    const formulasToWrite = [];
    const rowsToHighlight = [];

    dataRows.forEach(({ row, originalIndex }) => {
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

        // --- LDA Follow-up Logic ---
        if (ldaFollowUpMap && studentNumberColIdx !== -1) {
            const studentId = row[studentNumberColIdx];
            if (studentId && ldaFollowUpMap.has(studentId)) {
                const followUp = ldaFollowUpMap.get(studentId);
                const outreachColInFinal = findColumnIndex(finalHeaders.map(h => h.toLowerCase()), CONSTANTS.OUTREACH_COLS);
                if (outreachColInFinal !== -1) {
                    const formattedDate = `${followUp.date.getMonth() + 1}-${followUp.date.getDate()}-${String(followUp.date.getFullYear()).slice(-2)}`;
                    newRow[outreachColInFinal] = `[Retention Add-In] Will Engage ${formattedDate}`;
                    rowsToHighlight.push({ rowIndex: dataToWrite.length, color: "#FFEDD5" }); // Light Orange
                }
            }
        }
        // --- End LDA Follow-up Logic ---

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
        
        await context.sync();

        // --- DNC Formatting Logic ---
        if (dncStatusMap && dncStatusMap.size > 0) {
            console.log("[DEBUG] Applying DNC formatting to table:", tableName);
            const tableBodyRange = table.getDataBodyRange();
            const finalLCHeaders = finalHeaders.map(h => h.toLowerCase());
            const studentIdColInTable = findColumnIndex(finalLCHeaders, CONSTANTS.STUDENT_NUMBER_COLS);
            
            if (studentIdColInTable !== -1) {
                tableBodyRange.load("values");
                await context.sync();
                
                for (let i = 0; i < tableBodyRange.values.length; i++) {
                    const studentId = String(tableBodyRange.values[i][studentIdColInTable]);
                    if (studentId && dncStatusMap.has(studentId)) {
                        const dncInfo = dncStatusMap.get(studentId);
                        const contactCols = {
                            phone: ["primary phone", "other phone"],
                            email: ["student email", "personal email"]
                        };
                        
                        const applyDncFormat = (colName) => {
                            const colIdx = finalLCHeaders.indexOf(colName);
                            if (colIdx !== -1) {
                                const cell = tableBodyRange.getCell(i, colIdx);
                                cell.format.fill.color = "#C00000"; // Dark Red
                                cell.format.font.strikethrough = true;
                            }
                        };

                        if (dncInfo.all) {
                            [...contactCols.phone, ...contactCols.email].forEach(applyDncFormat);
                            const outreachColIdx = finalLCHeaders.indexOf("outreach");
                            if (outreachColIdx !== -1) {
                                const cell = tableBodyRange.getCell(i, outreachColIdx);
                                cell.values = [["Student wishes not to be contacted."]];
                            }
                        } else {
                            if (dncInfo.phone) contactCols.phone.forEach(applyDncFormat);
                            if (dncInfo.email) contactCols.email.forEach(applyDncFormat);
                        }
                    }
                }
            }
        }
        // --- End DNC Formatting ---

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
        
        // Apply LDA Follow-up Highlights to the specified range
        if (rowsToHighlight.length > 0) {
            console.log(`[DEBUG] Applying ${rowsToHighlight.length} LDA follow-up highlights.`);
            const finalLCHeaders = finalHeaders.map(h => h.toLowerCase());
            const studentNameColIdxInTable = findColumnIndex(finalLCHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const outreachColIdxInTable = findColumnIndex(finalLCHeaders, CONSTANTS.OUTREACH_COLS);

            if (studentNameColIdxInTable !== -1 && outreachColIdxInTable !== -1) {
                const startCol = Math.min(studentNameColIdxInTable, outreachColIdxInTable);
                const colCount = Math.abs(studentNameColIdxInTable - outreachColIdxInTable) + 1;

                for (const item of rowsToHighlight) {
                    const rowRange = table.getDataBodyRange().getRow(item.rowIndex);
                    const highlightRange = rowRange.getCell(0, startCol).getResizedRange(0, colCount - 1);
                    highlightRange.format.fill.color = item.color;
                }
            } else {
                console.warn("[DEBUG] Could not find 'StudentName' and/or 'Outreach' columns in the final table to apply highlights.");
                // Fallback to highlighting the whole row if columns aren't found
                for (const item of rowsToHighlight) {
                    const rowRange = table.getDataBodyRange().getRow(item.rowIndex);
                    rowRange.format.fill.color = item.color;
                }
            }
        }

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
                        const columnLetter = String.fromCharCode(65 + idx);
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
