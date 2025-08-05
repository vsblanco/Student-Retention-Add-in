/*
 * This file contains the logic for the "Import Data" ribbon button command.
 */
import { CONSTANTS, errorHandler, parseDate, jsDateToExcelDate, normalizeName, formatToLastFirst, dataUrlToArrayBuffer, parseCsvRow, findColumnIndex, getSettings } from './utils.js';

let importDialog = null;

/**
 * Sends a status message back to the import dialog.
 * @param {string} status The message to send.
 * @param {string} type The type of message ('log', 'error', 'complete').
 * @param {string[]} [details] Optional array of strings for collapsible details.
 */
function sendMessageToDialog(status, type = 'log', details = []) {
    if (importDialog) {
        console.log(`[DIALOG LOG] ${type.toUpperCase()}: ${status}`);
        if (details.length > 0) {
             console.log(details.map(d => `  - ${d}`).join('\n'));
        }
        importDialog.messageChild(JSON.stringify({ type, status, details }));
    }
}


/**
 * Opens a dialog for data import.
 * @param {Office.AddinCommands.Event} event The event object passed from the ribbon button.
 */
export function openImportDialog(event) {
    Office.context.ui.displayDialogAsync(
        'https://vsblanco.github.io/Student-Retention-Add-in/commands/import-dialog.html',
        { height: 55, width: 45, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog failed to open: " + asyncResult.error.message);
                event.completed();
                return;
            }
            importDialog = asyncResult.value;
            importDialog.addEventHandler(Office.EventType.DialogMessageReceived, processImportMessage);
            event.completed();
        }
    );
}

/**
 * Routes messages from the dialog to the appropriate handler.
 * @param {Office.DialogMessageReceivedEventArgs} arg The event args.
 */
async function processImportMessage(arg) {
    const message = JSON.parse(arg.message);

    switch (message.type) {
        case 'dialogReady':
            await handleDialogReady();
            break;
        case 'createMasterList':
            await handleCreateMasterList(message.columns);
            break;
        case 'fileSelected':
            await handleFileSelected(message);
            break;
        case 'updateMaster':
            await handleUpdateMaster(message);
            break;
        case 'updateGrades':
            await handleUpdateGrades(message);
            break;
        case 'closeDialog':
            if (importDialog) {
                importDialog.close();
            }
            break;
        default:
            console.error("Unknown message type from dialog:", message.type);
            if (importDialog) {
                importDialog.close();
            }
    }
}

/**
 * Checks if the "Master List" sheet exists and tells the dialog which UI to show.
 */
async function handleDialogReady() {
    try {
        await Excel.run(async (context) => {
            const sheetNames = context.workbook.worksheets.load("items/name");
            await context.sync();
            let masterListExists = false;
            for (let i = 0; i < sheetNames.items.length; i++) {
                if (sheetNames.items[i].name === CONSTANTS.MASTER_LIST_SHEET) {
                    masterListExists = true;
                    break;
                }
            }
            if (importDialog) {
                importDialog.messageChild(JSON.stringify({ type: 'masterListCheckResult', exists: masterListExists }));
            }
        });
    } catch (error) {
        console.error("Error checking for Master List:", error);
        if (importDialog) {
            importDialog.messageChild(JSON.stringify({ type: 'masterListCheckResult', exists: false }));
        }
    }
}

/**
 * Creates the "Master List" sheet with the specified columns.
 * @param {string[]} columns An array of column header names.
 */
async function handleCreateMasterList(columns) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.add(CONSTANTS.MASTER_LIST_SHEET);
            const headerRange = sheet.getRangeByIndexes(0, 0, 1, columns.length);
            headerRange.values = [columns];
            headerRange.format.font.bold = true;
            sheet.getUsedRange().format.autofitColumns();
            sheet.activate();
            await context.sync();
        });
        if (importDialog) {
            importDialog.messageChild(JSON.stringify({ type: 'masterListCreated' }));
        }
    } catch (error) {
        console.error("Error creating Master List:", error);
        // We don't have a specific UI for this error in the dialog, but we could add one.
    }
}


/**
 * Handles the file selection event from the dialog.
 * Checks for required columns and sheet existence to enable/disable dialog buttons.
 * @param {object} message The message from the dialog.
 */
async function handleFileSelected(message) {
    const { fileName, data: dataUrl } = message;
    sendMessageToDialog(`File selected: ${fileName}. Starting analysis...`);
    let hasMasterListSheet = false;

    try {
        const arrayBuffer = dataUrlToArrayBuffer(dataUrl);
        const workbook = new ExcelJS.Workbook();
        
        if (fileName.toLowerCase().endsWith('.xlsx')) {
            sendMessageToDialog("Detected .xlsx file, parsing...");
            await workbook.xlsx.load(arrayBuffer);
        } else {
            sendMessageToDialog("Detected .csv file, parsing...");
            const csvData = new TextDecoder("utf-8").decode(arrayBuffer);
            const rows = csvData.split(/\r?\n/).filter(row => row.trim().length > 0);
            const data = rows.map(row => parseCsvRow(row));
            const worksheet = workbook.addWorksheet('sheet1');
            worksheet.addRows(data);
        }
        sendMessageToDialog("File parsed successfully. Reading headers...");

        const worksheet = workbook.worksheets[0];
        const headers = [];
        worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell) => {
            headers.push(String(cell.value || ''));
        });
        const lowerCaseHeaders = headers.map(h => h.toLowerCase());
        sendMessageToDialog(`Found headers: [${headers.join(', ')}]`);
        
        // --- Column Checks ---
        sendMessageToDialog("Checking for required columns...");
        const checkCol = (key, aliases) => {
            const index = findColumnIndex(lowerCaseHeaders, aliases);
            sendMessageToDialog(`- Checking for '${key}': ${index !== -1 ? `Found at index ${index}` : 'Not found'}`);
            return index !== -1;
        };

        const hasStudentNameCol = checkCol('Student Name', CONSTANTS.STUDENT_NAME_COLS);
        const hasStudentIdCol = checkCol('Student ID', CONSTANTS.STUDENT_ID_COLS);
        const hasStudentNumberCol = checkCol('Student Number', CONSTANTS.STUDENT_NUMBER_COLS);
        const hasCourseCol = checkCol('Course', CONSTANTS.COLUMN_MAPPINGS.course);
        const hasCourseIdCol = checkCol('Course ID', CONSTANTS.COLUMN_MAPPINGS.courseId);
        const hasCourseLastAccessCol = checkCol('Course Last Access', CONSTANTS.COLUMN_MAPPINGS.courseLastAccess);
        const hasCurrentScoreCol = checkCol('Current Score', CONSTANTS.COLUMN_MAPPINGS.currentScore);
        
        // --- Sheet Existence Check ---
        if (hasStudentNameCol || hasStudentNumberCol) {
            sendMessageToDialog("Checking for 'Master List' sheet in workbook...");
            await Excel.run(async (context) => {
                const sheetNames = context.workbook.worksheets.load("items/name");
                await context.sync();
                for (let i = 0; i < sheetNames.items.length; i++) {
                    if (sheetNames.items[i].name === CONSTANTS.MASTER_LIST_SHEET) {
                        hasMasterListSheet = true;
                        break;
                    }
                }
            });
            sendMessageToDialog(`'Master List' sheet exists: ${hasMasterListSheet}`);
        }
        
        // --- Determine Available Actions ---
        const isGradeFile = hasStudentNameCol && hasCurrentScoreCol && hasCourseCol && (hasCourseIdCol || hasCourseLastAccessCol);
        const isMasterFile = hasStudentNumberCol && !isGradeFile;
        sendMessageToDialog(`File type analysis: Grade File=${isGradeFile}, Master File=${isMasterFile}`);

        let statusMessage = "";
        let canUpdateGrades = false;
        let canUpdateMaster = false;

        if (!hasMasterListSheet) {
            statusMessage = "'Master List' sheet not found in this workbook.";
            sendMessageToDialog(statusMessage, 'error');
        } else if (isGradeFile) {
            statusMessage = "Ready to update grades.";
            canUpdateGrades = true;
        } else if (isMasterFile) {
            statusMessage = "Ready to update Master List.";
            canUpdateMaster = true;
        } else {
            statusMessage = "File not compatible. Check log for missing columns.";
            sendMessageToDialog("File does not meet requirements for either update type.", 'error');
        }
        
        if (importDialog) {
            importDialog.messageChild(JSON.stringify({ 
                canUpdateMaster: canUpdateMaster,
                canUpdateGrades: canUpdateGrades,
                status: statusMessage
            }));
        }
    } catch (error) {
        console.error("Error during file check:", error);
        sendMessageToDialog(`Error during file analysis: ${error.message}`, 'error');
        if (importDialog) {
            importDialog.messageChild(JSON.stringify({ 
                canUpdateMaster: false,
                canUpdateGrades: false,
                status: `Error: ${error.message}`
            }));
        }
    }
}


/**
 * Handles the Master List update action. It clears all data rows, then performs a clean import,
 * highlighting any students that were not on the list before the import.
 * @param {object} message The message from the dialog.
 */
async function handleUpdateMaster(message) {
    sendMessageToDialog("Starting Master List update process...");
    try {
        // 1. Parse user's uploaded file
        sendMessageToDialog("Parsing uploaded file...");
        const userArrayBuffer = dataUrlToArrayBuffer(message.data);
        const userWorkbook = new ExcelJS.Workbook();
        if (message.fileName.toLowerCase().endsWith('.xlsx')) {
            await userWorkbook.xlsx.load(userArrayBuffer);
        } else {
            const csvData = new TextDecoder("utf-8").decode(userArrayBuffer);
            const rows = csvData.split(/\r?\n/).filter(row => row.trim().length > 0);
            const data = rows.map(row => parseCsvRow(row));
            const worksheet = userWorkbook.addWorksheet('sheet1');
            worksheet.addRows(data);
        }
        const userWorksheet = userWorkbook.worksheets[0];

        const userHeaders = [];
        userWorksheet.getRow(1).eachCell({ includeEmpty: true }, (cell) => {
            userHeaders.push(String(cell.value || ''));
        });
        const lowerCaseUserHeaders = userHeaders.map(h => h.toLowerCase());
        sendMessageToDialog(`Parsed import headers: [${userHeaders.join(', ')}]`);

        const userData = [];
        userWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const rowData = [];
                row.eachCell({ includeEmpty: true }, (cell) => {
                    rowData.push(cell.value);
                });
                userData.push(rowData);
            }
        });
        sendMessageToDialog(`Parsed ${userData.length} data rows from imported file.`);

        const userStudentNameCol = findColumnIndex(lowerCaseUserHeaders, CONSTANTS.STUDENT_NAME_COLS);
        if (userStudentNameCol === -1) {
            throw new Error("Imported file is missing a 'Student Name' column.");
        }

        // 2. Read existing Master List to identify new students before clearing
        let newStudents = [];
        let existingStudents = [];
        let masterHeaders;
        let lowerCaseMasterHeaders;
        let masterStudentNameCol;
        let colMapping;

        await Excel.run(async (context) => {
            sendMessageToDialog("Reading current 'Master List' to identify new students...");
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            if (usedRange.values.length < 1) {
                throw new Error("'Master List' is empty or has no header row.");
            }

            masterHeaders = usedRange.values[0].map(h => String(h || ''));
            lowerCaseMasterHeaders = masterHeaders.map(h => h.toLowerCase());
            masterStudentNameCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            sendMessageToDialog(`'Master List' headers: [${masterHeaders.join(', ')}]`);

            if (masterStudentNameCol === -1) {
                throw new Error("'Master List' is missing a 'StudentName' column.");
            }

            const masterNameMap = new Map();
            for (let i = 1; i < usedRange.values.length; i++) {
                const name = usedRange.values[i][masterStudentNameCol];
                if (name) {
                    masterNameMap.set(normalizeName(name), i);
                }
            }
            sendMessageToDialog(`Created map of ${masterNameMap.size} students from 'Master List'.`);

            colMapping = lowerCaseUserHeaders.map(userHeader =>
                lowerCaseMasterHeaders.indexOf(userHeader)
            );
            sendMessageToDialog(`Column mapping created: [${colMapping.join(', ')}]`);

            for (const userRow of userData) {
                const studentName = userRow[userStudentNameCol];
                const normalizedName = normalizeName(studentName);

                if (masterNameMap.has(normalizedName)) {
                    existingStudents.push(userRow);
                } else {
                    newStudents.push(userRow);
                }
            }
            sendMessageToDialog(`Found ${newStudents.length} new students and ${existingStudents.length} existing students.`);
        });

        // 3. Clear the sheet and repopulate
        sendMessageToDialog("Clearing 'Master List' for a clean import...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();

            if (usedRange.rowCount > 1) {
                const rangeToClear = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, masterHeaders.length);
                rangeToClear.clear(Excel.ClearApplyTo.all);
                rangeToClear.getEntireRow().delete(Excel.DeleteShiftDirection.up);
            }
            await context.sync();
            sendMessageToDialog("Sheet cleared.");

            // 4. Combine students and prepare data for writing
            const allStudentsToWrite = [...newStudents, ...existingStudents];
            if (allStudentsToWrite.length === 0) {
                sendMessageToDialog("No students to import.", 'complete');
                return;
            }

            sendMessageToDialog(`Preparing to write ${allStudentsToWrite.length} students...`);
            const dataToWrite = allStudentsToWrite.map(userRow => {
                const newRow = new Array(masterHeaders.length).fill("");
                for (let userColIdx = 0; userColIdx < userRow.length; userColIdx++) {
                    const masterColIdx = colMapping[userColIdx];
                    if (masterColIdx !== -1) {
                        let cellValue = userRow[userColIdx] || "";
                        if (masterColIdx === masterStudentNameCol) {
                            cellValue = formatToLastFirst(String(cellValue));
                        }
                        newRow[masterColIdx] = cellValue;
                    }
                }
                
                // *** FIX START: Timezone bug fix for LDA dates ***
                const userLdaColIdx = findColumnIndex(lowerCaseUserHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
                if (userLdaColIdx !== -1) {
                    const ldaValue = userRow[userLdaColIdx];
                    const ldaDate = parseDate(ldaValue);
                    if (ldaDate) {
                        // Get today's date at midnight UTC to ensure correct "days out" calculation
                        const today = new Date();
                        const todayUTC = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
                        
                        // The ldaDate from parseDate is already a UTC date object (representing midnight UTC).
                        const masterLdaColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
                        const masterDaysOutColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);

                        if (masterLdaColIdx !== -1) {
                            // jsDateToExcelDate uses getTime(), which is UTC based, so this is correct.
                            newRow[masterLdaColIdx] = jsDateToExcelDate(ldaDate);
                        }
                        if (masterDaysOutColIdx !== -1) {
                            // Calculate difference in days based on UTC timestamps to avoid timezone errors.
                            const daysOut = Math.floor((todayUTC.getTime() - ldaDate.getTime()) / (1000 * 60 * 60 * 24));
                            newRow[masterDaysOutColIdx] = daysOut >= 0 ? daysOut : 0; // Ensure days out is not negative
                        }
                    }
                }
                // *** FIX END ***
                return newRow;
            });

            // 5. Write all data in a single batch
            sendMessageToDialog("Writing data to the sheet...");
            const writeRange = sheet.getRangeByIndexes(1, 0, dataToWrite.length, masterHeaders.length);
            writeRange.values = dataToWrite;
            await context.sync();

            // 6. Highlight new students
            if (newStudents.length > 0) {
                sendMessageToDialog(`Highlighting ${newStudents.length} new students...`);
                const highlightRange = sheet.getRangeByIndexes(1, 0, newStudents.length, masterHeaders.length);
                highlightRange.format.fill.color = "#ADD8E6"; // Light Blue
            }
            
            // 7. Final formatting and autofit
            const masterLdaColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
            if (masterLdaColIdx !== -1) {
                sendMessageToDialog("Formatting 'LDA' column as date...");
                const ldaColumn = sheet.getRangeByIndexes(0, masterLdaColIdx, dataToWrite.length + 1, 1);
                ldaColumn.numberFormat = [["M-DD-YYYY"]];
            }

            sendMessageToDialog("Autofitting columns for readability...");
            sheet.getUsedRange().format.autofitColumns();
            
            await context.sync();
            sendMessageToDialog("Master List update process completed successfully.", 'complete');
        });

    } catch (error) {
        console.error("Error updating Master List: " + error);
        sendMessageToDialog(`Error: ${error.message}`, 'error');
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}


/**
 * Handles updating grades and gradebook links in the Master List for existing students.
 * This version is optimized to perform a single bulk write operation to avoid timeouts.
 * @param {object} message The message from the dialog.
 */
async function handleUpdateGrades(message) {
    sendMessageToDialog("Starting grade update process...");
    try {
        const settings = await getSettings();
        const { treatEmptyGradesAsZero } = settings.createlda;
        sendMessageToDialog(`Setting 'Treat Empty Grades as 0' is ${treatEmptyGradesAsZero ? 'ON' : 'OFF'}.`);

        // Step 1: Parse the uploaded file using ExcelJS
        sendMessageToDialog("Parsing uploaded file for grade data...");
        const userArrayBuffer = dataUrlToArrayBuffer(message.data);
        const userWorkbook = new ExcelJS.Workbook();
        if (message.fileName.toLowerCase().endsWith('.xlsx')) {
            await userWorkbook.xlsx.load(userArrayBuffer);
        } else {
            const csvData = new TextDecoder("utf-8").decode(userArrayBuffer);
            const rows = csvData.split(/\r?\n/).filter(row => row.trim().length > 0);
            const data = rows.map(row => parseCsvRow(row));
            const worksheet = userWorkbook.addWorksheet('sheet1');
            worksheet.addRows(data);
        }
        const userWorksheet = userWorkbook.worksheets[0];
        
        const userHeaders = [];
        userWorksheet.getRow(1).eachCell({ includeEmpty: true }, (cell) => {
            userHeaders.push(String(cell.value || '').toLowerCase());
        });
        sendMessageToDialog(`Parsed grade file headers: [${userHeaders.join(', ')}]`);
        
        const userData = [];
        userWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const rowData = [];
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    rowData[colNumber - 1] = cell.value;
                });
                userData.push(rowData);
            }
        });
        sendMessageToDialog(`Parsed ${userData.length} rows from the grade file.`);

        // Step 2: Find column indices in the imported file
        const userStudentNameCol = findColumnIndex(userHeaders, CONSTANTS.STUDENT_NAME_COLS);
        const userStudentIdCol = findColumnIndex(userHeaders, CONSTANTS.STUDENT_ID_COLS);
        const userCourseIdCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.courseId);
        const userGradeCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.currentScore);
        const userCourseCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.course);
        const userMissingAssignmentsCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
        const userZeroAssignmentsCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.courseZeroAssignments);

        if (userGradeCol === -1 || userStudentNameCol === -1 || userCourseCol === -1) {
            throw new Error("Imported file is missing one of the required columns: Student Name, Course, or Current Score/Grade.");
        }
        sendMessageToDialog("Required columns found in grade file.");
        sendMessageToDialog(`- Using column index ${userGradeCol} for Grades.`);
        if (userMissingAssignmentsCol !== -1) {
            sendMessageToDialog(`- Using column index ${userMissingAssignmentsCol} for Missing Assignments.`);
        }
        if (userZeroAssignmentsCol !== -1) {
            sendMessageToDialog(`- Using column index ${userZeroAssignmentsCol} for Zero Assignments.`);
        }
        if (userCourseIdCol !== -1 && userStudentIdCol !== -1) {
            sendMessageToDialog(`- Using column indices ${userCourseIdCol} (CourseID) and ${userStudentIdCol} (StudentID) for hyperlinks.`);
        }


        // Step 3: Create a map of student data from the import, filtering out 'CAPV' courses
        const studentDataMap = new Map();
        let skippedRows = 0;
        userData.forEach(row => {
            const courseName = row[userCourseCol] ? String(row[userCourseCol]) : '';
            if (courseName.toUpperCase().includes('CAPV')) {
                skippedRows++;
                return; // Skip this row
            }

            const studentName = row[userStudentNameCol];
            if (studentName) {
                const normalized = normalizeName(studentName);
                let gradeValue = row[userGradeCol];
                if ((gradeValue === null || gradeValue === undefined || gradeValue === '') && treatEmptyGradesAsZero) {
                    gradeValue = 0;
                }
                studentDataMap.set(normalized, {
                    grade: gradeValue,
                    courseId: row[userCourseIdCol],
                    studentId: row[userStudentIdCol],
                    missingAssignments: row[userMissingAssignmentsCol],
                    zeroAssignments: row[userZeroAssignmentsCol]
                });
            }
        });
        sendMessageToDialog(`Created a map of ${studentDataMap.size} students (skipped ${skippedRows} 'CAPV' rows).`);

        // Step 4: Perform the bulk update on the "Master List" sheet
        await Excel.run(async (context) => {
            sendMessageToDialog("Reading current 'Master List' data and formulas into memory...");
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const range = sheet.getUsedRange();
            
            // Load both values and formulas from the sheet into memory
            range.load("values, formulas, rowCount");
            await context.sync();
            sendMessageToDialog("'Master List' data and formulas loaded.");

            const masterFormulas = range.formulas;
            // Create a mutable copy of the formulas to preserve existing data/formulas
            const valuesToWrite = masterFormulas.map(row => [...row]);

            const masterHeaders = range.values[0].map(h => String(h || '').toLowerCase());

            // Find column indices in the Master List
            const masterStudentNameCol = findColumnIndex(masterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterGradeCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);
            const masterGradebookCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            const masterMissingAssignmentsCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const masterZeroAssignmentsCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseZeroAssignments);
            
            if (masterStudentNameCol === -1 || masterGradeCol === -1 || masterGradebookCol === -1) {
                throw new Error("'Master List' is missing required columns: StudentName, Grade, or Grade Book.");
            }
            sendMessageToDialog("Required columns located in 'Master List'.");

            let updatedCount = 0;
            let gradesUpdated = 0;
            let missingUpdated = 0;
            let zerosUpdated = 0;
            let linksUpdated = 0;

            // Iterate through the Master List data (in memory) and update the valuesToWrite array
            for (let i = 1; i < range.rowCount; i++) {
                const masterName = range.values[i][masterStudentNameCol]; // Use original values for matching
                if (masterName) {
                    const normalizedName = normalizeName(masterName);
                    if (studentDataMap.has(normalizedName)) {
                        const importedData = studentDataMap.get(normalizedName);
                        const details = [];
                        
                        // Update grade
                        if (importedData.grade !== undefined && importedData.grade !== null) {
                            valuesToWrite[i][masterGradeCol] = importedData.grade;
                            details.push(`- Grade set to: ${importedData.grade}`);
                            gradesUpdated++;
                        }

                        // Update assignments
                        if (masterMissingAssignmentsCol !== -1 && importedData.missingAssignments !== undefined) {
                            valuesToWrite[i][masterMissingAssignmentsCol] = importedData.missingAssignments;
                            details.push(`- Missing Assignments set to: ${importedData.missingAssignments}`);
                            missingUpdated++;
                        }
                        if (masterZeroAssignmentsCol !== -1 && importedData.zeroAssignments !== undefined) {
                            valuesToWrite[i][masterZeroAssignmentsCol] = importedData.zeroAssignments;
                            details.push(`- Zero Assignments set to: ${importedData.zeroAssignments}`);
                            zerosUpdated++;
                        }

                        // Update hyperlink by setting a formula string in the values array
                        if (importedData.courseId && importedData.studentId) {
                            details.push(`- Found CourseID: ${importedData.courseId}, StudentID: ${importedData.studentId}`);
                            const newGradebookLink = `https://nuc.instructure.com/courses/${importedData.courseId}/grades/${importedData.studentId}`;
                            valuesToWrite[i][masterGradebookCol] = `=HYPERLINK("${newGradebookLink}", "Gradebook")`;
                            details.push(`- Wrapped hyperlink: ${newGradebookLink}`);
                            linksUpdated++;
                        }
                        
                        if(details.length > 0) {
                            sendMessageToDialog(`Updating row ${i + 1}: ${masterName}`, 'log', details);
                        }
                        updatedCount++;
                    }
                }
            }
            
            sendMessageToDialog(`Prepared updates for ${updatedCount} students:`);
            sendMessageToDialog(`- ${gradesUpdated} grades to be updated.`);
            if (missingUpdated > 0) sendMessageToDialog(`- ${missingUpdated} 'Missing Assignments' values to be updated.`);
            if (zerosUpdated > 0) sendMessageToDialog(`- ${zerosUpdated} 'Zero Assignments' values to be updated.`);
            if (linksUpdated > 0) sendMessageToDialog(`- ${linksUpdated} hyperlinks to be created/updated.`);


            if (updatedCount > 0) {
                // Write the updated data back to the sheet in a single operation.
                // Excel will automatically interpret values starting with '=' as formulas.
                sendMessageToDialog("Writing all updates to the sheet at once...");
                range.values = valuesToWrite;
                sheet.getUsedRange().format.autofitColumns();
            }
            
            await context.sync();

            // Apply conditional formatting after updates are synced
            await applyGradeConditionalFormatting(context);

            sendMessageToDialog("Grade update process completed successfully.", 'complete');
        });

    } catch (error) {
        console.error("Error updating grades in Master List: " + error);
        sendMessageToDialog(`Error: ${error.message}`, 'error');
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Applies a 3-color scale conditional formatting to the grade column.
 * @param {Excel.RequestContext} context The request context.
 */
async function applyGradeConditionalFormatting(context) {
    sendMessageToDialog("Applying conditional formatting to grades...");
    const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
    const range = sheet.getUsedRange(true); // Use true to get the most up-to-date range
    range.load("values, rowCount");
    await context.sync();

    const headers = range.values[0].map(h => String(h || '').toLowerCase());
    const gradeColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.grade);

    if (gradeColIdx === -1) {
        sendMessageToDialog("'Grade' column not found, skipping conditional formatting.", 'log');
        return;
    }

    if (range.rowCount <= 1) {
        sendMessageToDialog("No data rows to format.", 'log');
        return;
    }

    const gradeColumnRange = sheet.getRangeByIndexes(1, gradeColIdx, range.rowCount - 1, 1);

    // Determine if grades are 0-1 or 0-100 scale by checking the first few values
    let isPercentScale = false;
    for (let i = 1; i < Math.min(range.rowCount, 10); i++) { // Check up to 10 rows
        if (range.values[i] && typeof range.values[i][gradeColIdx] === 'number' && range.values[i][gradeColIdx] > 1) {
            isPercentScale = true;
            break;
        }
    }
    
    sendMessageToDialog(`Detected grade scale: ${isPercentScale ? '0-100' : '0-1'}. Applying 3-color scale.`);

    // Clear existing conditional formats on the column to avoid duplicates
    gradeColumnRange.conditionalFormats.clearAll();
    
    const conditionalFormat = gradeColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    const criteria = {
        minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" }, // Red
        midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: isPercentScale ? "70" : "0.7", color: "#FFEB84" }, // Yellow
        maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" } // Green
    };
    conditionalFormat.colorScale.criteria = criteria;

    await context.sync();
    sendMessageToDialog("Conditional formatting applied.");
}
