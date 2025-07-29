/*
 * This file contains the logic for ribbon button commands.
 * It is loaded by commands.html, which is specified as the FunctionFile in the manifest.
 */

const CONSTANTS = {
    // NOTE: "Student ID" and "Student Number" are treated as distinct values.
    // "Student ID" refers to the Canvas ID, used for creating gradebook links.
    // "Student Number" (and "Student Identifier") refers to the internal school ID.
    STUDENT_NAME_COLS: ["studentname", "student name","student"],
    OUTREACH_COLS: ["outreach"],
    STUDENT_ID_COLS: ["student id"],
    STUDENT_NUMBER_COLS: ["studentnumber", "student identifier"],
    MASTER_LIST_SHEET: "Master List",
    TEMPLATE_URL: 'https://vsblanco.github.io/Student-Retention-Add-in/Template.xlsx',
    SETTINGS_KEY: "studentRetentionSettings", // Key for document settings
    COLUMN_MAPPINGS: {
        course: ["course"],
        courseId: ["course id"],
        courseLastAccess: ["course last access"],
        currentScore: ["current score", "grade", "course grade"],
        grade: ["grade", "course grade"],
        gradeBook: ["grade book", "gradebook"],
        daysOut: ["days out"],
        lastLda: ["lda", "last lda"],
        assigned: ["assigned"],
        programVersion: ["programversion", "program version"],
        courseMissingAssignments: ["course missing assignments"],
        courseZeroAssignments: ["course zero assignments"]
    }
};

let importDialog = null;
let transferDialog = null;
let createLdaDialog = null; // New dialog variable

/**
 * Gets the settings object from document settings, ensuring it's the latest version.
 * @returns {Promise<object>} A promise that resolves with the parsed settings object.
 */
async function getSettings() {
    // First, refresh the settings from the document to ensure we have the latest version.
    await new Promise((resolve) => {
        Office.context.document.settings.refreshAsync(asyncResult => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error refreshing settings: " + asyncResult.error.message);
                // Even if refresh fails, we proceed with the cached version.
            } else {
                console.log("Settings refreshed successfully.");
            }
            resolve();
        });
    });

    // Now, get the settings value.
    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEY);
    const defaults = {
        createlda: {
            daysOutFilter: 6,
            includeFailingList: true,
            hideLeftoverColumns: true,
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        }
    };

    if (settingsString) {
        try {
            const settings = JSON.parse(settingsString);
            settings.createlda = { ...defaults.createlda, ...(settings.createlda || {}) };
            return settings;
        } catch (e) {
            console.error("Error parsing settings, returning defaults:", e);
            return defaults;
        }
    }
    return defaults;
}


/**
 * Parses a date value from various possible formats (Date object, string, Excel serial number).
 * @param {*} dateValue The value to parse.
 * @returns {Date|null} A valid Date object or null.
 */
function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
        // ExcelJS can return Date objects for .xlsx
        return dateValue;
    }
    if (typeof dateValue === 'number') {
        // Excel serial date number
        // Check for a reasonable range to avoid treating random numbers as dates
        if (dateValue > 25569) { // Corresponds to 1970-01-01
            return new Date((dateValue - 25569) * 86400 * 1000);
        }
    }
    if (typeof dateValue === 'string') {
        // Try parsing common date formats
        const parsed = new Date(dateValue);
        if (!isNaN(parsed.getTime())) {
            return parsed;
        }
    }
    return null;
}

/**
 * Converts a JavaScript Date object to an Excel serial date number.
 * @param {Date} date The JavaScript Date object.
 * @returns {number} The Excel serial date number.
 */
function jsDateToExcelDate(date) {
    return (date.getTime() / 86400000) + 25569;
}


/**
 * Sends a status message back to the import dialog.
 * @param {string} status The message to send.
 * @param {string} type The type of message ('log', 'error', 'complete').
 */
function sendMessageToDialog(status, type = 'log') {
    if (importDialog) {
        console.log(`[DIALOG LOG] ${status}`);
        importDialog.messageChild(JSON.stringify({ type, status }));
    }
}

/**
 * Helper to normalize names from "Last, First" or "First Last" to "first last"
 * for consistent matching.
 * @param {string} name The name to normalize.
 * @returns {string} The normalized name.
 */
const normalizeName = (name) => {
    if (!name || typeof name !== 'string') return '';
    name = name.trim().toLowerCase();
    if (name.includes(',')) {
        const parts = name.split(',').map(part => part.trim());
        if (parts.length > 1) {
            return `${parts[1]} ${parts[0]}`;
        }
    }
    return name;
};

/**
 * Helper to format names to "Last, First" format.
 * @param {string} name The name to format.
 * @returns {string} The formatted name.
 */
const formatToLastFirst = (name) => {
    if (!name || typeof name !== 'string') return '';
    name = name.trim();
    if (name.includes(',')) {
        // Already "Last, First"
        return name.split(',').map(p => p.trim()).join(', ');
    }
    const parts = name.split(' ').filter(p => p);
    if (parts.length > 1) {
        const lastName = parts.pop();
        const firstName = parts.join(' ');
        return `${lastName}, ${firstName}`;
    }
    return name;
};


Office.onReady((info) => {
  // This file is loaded when the add-in is initialized, but the DOM is not relevant here.
});

/**
 * Opens a dialog for data import.
 * @param {Office.AddinCommands.Event} event The event object passed from the ribbon button.
 */
function openImportDialog(event) {
    Office.context.ui.displayDialogAsync(
        'https://vsblanco.github.io/Student-Retention-Add-in/commands/import-dialog.html',
        { height: 45, width: 35, displayInIframe: true },
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
 * Handles the file selection event from the dialog.
 * Checks for required columns and sheet existence to enable/disable dialog buttons.
 * @param {object} message The message from the dialog.
 */
async function handleFileSelected(message) {
    const { fileName, data: dataUrl } = message;
    console.log(`[DEBUG] File selected: ${fileName}`);
    let hasStudentIdCol = false;
    let hasStudentNumberCol = false;
    let hasMasterListSheet = false;
    let hasCourseCol = false;
    let hasCourseIdCol = false;
    let hasCourseLastAccessCol = false;
    let hasStudentNameCol = false;
    let hasCurrentScoreCol = false;

    try {
        const arrayBuffer = dataUrlToArrayBuffer(dataUrl);
        const workbook = new ExcelJS.Workbook();
        
        if (fileName.toLowerCase().endsWith('.xlsx')) {
            await workbook.xlsx.load(arrayBuffer);
        } else {
            const csvData = new TextDecoder("utf-8").decode(arrayBuffer);
            const rows = csvData.split(/\r?\n/).filter(row => row.trim().length > 0);
            const data = rows.map(row => parseCsvRow(row));
            const worksheet = workbook.addWorksheet('sheet1');
            worksheet.addRows(data);
        }
        console.log("[DEBUG] File parsed with ExcelJS.");

        const worksheet = workbook.worksheets[0];
        const headers = [];
        worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell) => {
            headers.push(String(cell.value || '').toLowerCase());
        });
        console.log("[DEBUG] Parsed headers:", headers);
        
        hasStudentIdCol = findColumnIndex(headers, CONSTANTS.STUDENT_ID_COLS) !== -1;
        hasStudentNumberCol = findColumnIndex(headers, CONSTANTS.STUDENT_NUMBER_COLS) !== -1;
        hasCourseCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.course) !== -1;
        hasCourseIdCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.courseId) !== -1;
        hasCourseLastAccessCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.courseLastAccess) !== -1;
        hasStudentNameCol = findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS) !== -1;
        hasCurrentScoreCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.currentScore) !== -1;
        console.log(`[DEBUG] Column checks: hasStudentIdCol=${hasStudentIdCol}, hasStudentNumberCol=${hasStudentNumberCol}, hasCourseCol=${hasCourseCol}, hasCourseIdCol=${hasCourseIdCol}, hasCourseLastAccessCol=${hasCourseLastAccessCol}, hasStudentNameCol=${hasStudentNameCol}, hasCurrentScoreCol=${hasCurrentScoreCol}`);

        if (hasStudentNameCol || hasStudentNumberCol) {
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
        }
        console.log(`[DEBUG] 'Master List' sheet exists: ${hasMasterListSheet}`);
        
        const isGradeFile = hasStudentNameCol && hasCurrentScoreCol && hasCourseCol && (hasCourseIdCol || hasCourseLastAccessCol);
        const isMasterFile = hasStudentNumberCol && !isGradeFile;
        console.log(`[DEBUG] File type detection: isGradeFile=${isGradeFile}, isMasterFile=${isMasterFile}`);

        let statusMessage = "";
        let canUpdateGrades = false;
        let canUpdateMaster = false;

        if (!hasMasterListSheet) {
            statusMessage = "'Master List' sheet not found in this workbook.";
        } else if (isGradeFile) {
            statusMessage = "Ready to update grades.";
            canUpdateGrades = true;
        } else if (isMasterFile) {
            statusMessage = "Ready to update Master List.";
            canUpdateMaster = true;
        } else {
            statusMessage = "File not compatible. Missing required columns.";
        }
        console.log(`[DEBUG] Final status: message='${statusMessage}', canUpdateGrades=${canUpdateGrades}, canUpdateMaster=${canUpdateMaster}`);

        if (importDialog) {
            importDialog.messageChild(JSON.stringify({ 
                canUpdateMaster: canUpdateMaster,
                canUpdateGrades: canUpdateGrades,
                status: statusMessage
            }));
        }
    } catch (error) {
        console.error("Error during file check:", error);
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
 * Handles the Master List update action. It updates existing students and adds new ones to the top, in batches.
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

        // 2. Initial data separation
        let newStudents = [];
        let existingStudents = [];
        let masterHeaders;
        let lowerCaseMasterHeaders;
        let masterStudentNameCol;
        let colMapping;

        await Excel.run(async (context) => {
            sendMessageToDialog("Accessing 'Master List' sheet for initial analysis...");
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();
            
            masterHeaders = usedRange.values[0].map(h => String(h || ''));
            lowerCaseMasterHeaders = masterHeaders.map(h => h.toLowerCase());
            masterStudentNameCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.STUDENT_NAME_COLS);

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

            colMapping = lowerCaseUserHeaders.map(userHeader =>
                lowerCaseMasterHeaders.indexOf(userHeader)
            );

            for (const userRow of userData) {
                const studentName = userRow[userStudentNameCol];
                const normalizedName = normalizeName(studentName);

                if (masterNameMap.has(normalizedName)) {
                    existingStudents.push({
                        userRow: userRow,
                        masterRowIndex: masterNameMap.get(normalizedName)
                    });
                } else {
                    newStudents.push(userRow);
                }
            }
            sendMessageToDialog(`Found ${existingStudents.length} existing students and ${newStudents.length} new students.`);
        });

        // 3. Prepare sheet by clearing formatting and inserting all new rows at once
        sendMessageToDialog("Preparing sheet for updates...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();

            if (usedRange.rowCount > 1) {
                const rangeToClear = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, masterHeaders.length);
                rangeToClear.format.fill.clear();
            }

            if (newStudents.length > 0) {
                const insertRange = sheet.getRangeByIndexes(1, 0, newStudents.length, 1);
                insertRange.getEntireRow().insert(Excel.InsertShiftDirection.down);
            }
            await context.sync();
        });
        sendMessageToDialog("Sheet prepared.");

        // 4. Batch-populate the newly created rows
        const batchSize = 100;
        if (newStudents.length > 0) {
            sendMessageToDialog(`Populating data for ${newStudents.length} new students in batches of ${batchSize}...`);
            for (let i = 0; i < newStudents.length; i += batchSize) {
                const batch = newStudents.slice(i, i + batchSize);
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
                    const startRow = 1 + i;
                    
                    const dataForBatch = batch.map(userRow => {
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
                        
                        // NEW: Calculate Days Out and format LDA date
                        const userLdaColIdx = findColumnIndex(lowerCaseUserHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
                        if (userLdaColIdx !== -1) {
                            const ldaValue = userRow[userLdaColIdx];
                            const ldaDate = parseDate(ldaValue);
                            if (ldaDate) {
                                const today = new Date();
                                today.setHours(0, 0, 0, 0);
                                ldaDate.setHours(0, 0, 0, 0);

                                const masterLdaColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
                                const masterDaysOutColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);

                                if (masterLdaColIdx !== -1) {
                                    newRow[masterLdaColIdx] = jsDateToExcelDate(ldaDate);
                                }
                                if (masterDaysOutColIdx !== -1) {
                                    const daysOut = Math.floor((today.getTime() - ldaDate.getTime()) / (1000 * 60 * 60 * 24));
                                    newRow[masterDaysOutColIdx] = daysOut;
                                }
                            }
                        }

                        return newRow;
                    });

                    const rangeToPopulate = sheet.getRangeByIndexes(startRow, 0, batch.length, masterHeaders.length);
                    rangeToPopulate.values = dataForBatch;
                    rangeToPopulate.format.fill.color = "#ADD8E6";
                    
                    await context.sync();
                });
                sendMessageToDialog(`Populated batch of ${batch.length} students. (${i + batch.length}/${newStudents.length})`);
            }
        }

        // 5. Batch-update existing students
        if (existingStudents.length > 0) {
            sendMessageToDialog(`Updating ${existingStudents.length} existing students in batches...`);
            
            let updatedMasterNameMap = new Map();
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
                const usedRange = sheet.getUsedRange();
                usedRange.load("values");
                await context.sync();
                for (let i = 1; i < usedRange.values.length; i++) {
                    const name = usedRange.values[i][masterStudentNameCol];
                    if (name) {
                        updatedMasterNameMap.set(normalizeName(name), i);
                    }
                }
            });

            for (let i = 0; i < existingStudents.length; i += batchSize) {
                const batch = existingStudents.slice(i, i + batchSize);
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
                    
                    for (const student of batch) {
                        const { userRow } = student;
                        const studentName = userRow[userStudentNameCol];
                        const normalizedName = normalizeName(studentName);
                        const newMasterRowIndex = updatedMasterNameMap.get(normalizedName);

                        if (newMasterRowIndex !== undefined) {
                            // Update all mapped columns first
                            for (let userColIdx = 0; userColIdx < userRow.length; userColIdx++) {
                                const masterColIdx = colMapping[userColIdx];
                                if (masterColIdx !== -1) {
                                    const cell = sheet.getCell(newMasterRowIndex, masterColIdx);
                                    cell.values = [[userRow[userColIdx] || ""]];
                                }
                            }
                            
                            // NEW: Calculate and update Days Out and LDA
                            const userLdaColIdx = findColumnIndex(lowerCaseUserHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
                            if (userLdaColIdx !== -1) {
                                const ldaValue = userRow[userLdaColIdx];
                                const ldaDate = parseDate(ldaValue);
                                if (ldaDate) {
                                    const today = new Date();
                                    today.setHours(0, 0, 0, 0);
                                    ldaDate.setHours(0, 0, 0, 0);

                                    const masterLdaColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
                                    const masterDaysOutColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);

                                    if (masterLdaColIdx !== -1) {
                                        const cell = sheet.getCell(newMasterRowIndex, masterLdaColIdx);
                                        cell.values = [[jsDateToExcelDate(ldaDate)]];
                                    }
                                    if (masterDaysOutColIdx !== -1) {
                                        const daysOut = Math.floor((today.getTime() - ldaDate.getTime()) / (1000 * 60 * 60 * 24));
                                        const cell = sheet.getCell(newMasterRowIndex, masterDaysOutColIdx);
                                        cell.values = [[daysOut]];
                                    }
                                }
                            }
                        }
                    }
                    await context.sync();
                });
                sendMessageToDialog(`Updated batch of ${batch.length} existing students. (${i + batch.length}/${existingStudents.length})`);
            }
        }
        
        // 6. Final formatting and autofit
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();

            // Format LDA column
            const masterLdaColIdx = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda);
            if (masterLdaColIdx !== -1) {
                const ldaColumn = sheet.getRangeByIndexes(0, masterLdaColIdx, usedRange.rowCount, 1);
                ldaColumn.numberFormat = [["M-DD-YYYY"]];
            }

            if (newStudents.length > 0 || existingStudents.length > 0) {
                sendMessageToDialog("Autofitting columns...");
                sheet.getUsedRange().format.autofitColumns();
            }
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
        // Step 1: Parse the uploaded file using ExcelJS
        sendMessageToDialog("Parsing uploaded file for grade update...");
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
        sendMessageToDialog(`Parsed ${userData.length} rows from the imported file.`);

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

        // Step 3: Create a map of student data from the import, filtering out 'CAPV' courses
        const studentDataMap = new Map();
        userData.forEach(row => {
            const courseName = row[userCourseCol] ? String(row[userCourseCol]) : '';
            if (courseName.toUpperCase().includes('CAPV')) {
                return; // Skip this row
            }

            const studentName = row[userStudentNameCol];
            if (studentName) {
                const normalized = normalizeName(studentName);
                studentDataMap.set(normalized, {
                    grade: row[userGradeCol],
                    courseId: row[userCourseIdCol],
                    studentId: row[userStudentIdCol],
                    missingAssignments: row[userMissingAssignmentsCol],
                    zeroAssignments: row[userZeroAssignmentsCol]
                });
            }
        });
        sendMessageToDialog(`Created a map of ${studentDataMap.size} students from the imported file after filtering.`);

        // Step 4: Perform the bulk update on the "Master List" sheet
        await Excel.run(async (context) => {
            sendMessageToDialog("Accessing 'Master List' sheet...");
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const range = sheet.getUsedRange();
            
            // Load all values and formulas from the sheet into memory
            range.load("values, formulas, rowCount");
            await context.sync();

            const masterValues = range.values;
            const masterFormulas = range.formulas;
            const masterHeaders = masterValues[0].map(h => String(h || '').toLowerCase());

            // Find column indices in the Master List
            const masterStudentNameCol = findColumnIndex(masterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterGradeCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);
            const masterGradebookCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            const masterMissingAssignmentsCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const masterZeroAssignmentsCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseZeroAssignments);
            
            if (masterStudentNameCol === -1 || masterGradeCol === -1 || masterGradebookCol === -1) {
                throw new Error("'Master List' is missing required columns: StudentName, Grade, or Grade Book.");
            }

            let updatedCount = 0;
            // Iterate through the Master List data (in memory) and update it
            for (let i = 1; i < range.rowCount; i++) {
                const masterName = masterValues[i][masterStudentNameCol];
                if (masterName) {
                    const normalizedName = normalizeName(masterName);
                    if (studentDataMap.has(normalizedName)) {
                        const importedData = studentDataMap.get(normalizedName);
                        
                        // Update the values in the local arrays
                        masterValues[i][masterGradeCol] = importedData.grade;

                        if (masterMissingAssignmentsCol !== -1 && importedData.missingAssignments !== undefined) {
                            masterValues[i][masterMissingAssignmentsCol] = importedData.missingAssignments;
                        }
                        if (masterZeroAssignmentsCol !== -1 && importedData.zeroAssignments !== undefined) {
                            masterValues[i][masterZeroAssignmentsCol] = importedData.zeroAssignments;
                        }

                        // Update the formula for the hyperlink
                        if (importedData.courseId && importedData.studentId) {
                            const newGradebookLink = `https://nuc.instructure.com/courses/${importedData.courseId}/grades/${importedData.studentId}`;
                            masterFormulas[i][masterGradebookCol] = `=HYPERLINK("${newGradebookLink}", "Gradebook")`;
                        }
                        
                        updatedCount++;
                    }
                }
            }
            
            sendMessageToDialog(`Found and prepared updates for ${updatedCount} matching students.`);

            if (updatedCount > 0) {
                // Write the updated data and formulas back to the sheet in two bulk operations
                sendMessageToDialog("Writing updates to the sheet...");
                range.values = masterValues;
                range.formulas = masterFormulas;
                sheet.getUsedRange().format.autofitColumns();
            }
            
            await context.sync();
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
 * Converts a data URL to an ArrayBuffer.
 * @param {string} dataUrl The data URL.
 * @returns {ArrayBuffer}
 */
function dataUrlToArrayBuffer(dataUrl) {
    const base64String = dataUrl.substring(dataUrl.indexOf(',') + 1);
    const binaryString = window.atob(base64String);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

/**
 * A robust CSV row parser that handles quoted fields.
 */
function parseCsvRow(row) {
    const cells = [];
    let inQuotes = false;
    let cell = '';
    for (let i = 0; i < row.length; i++) {
        const char = row[i];
        if (char === '"') {
            if (inQuotes && row[i+1] === '"') {
                cell += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            cells.push(cell);
            cell = '';
        } else {
            cell += char;
        }
    }
    cells.push(cell);
    return cells;
}

/**
 * Finds the index of a column by checking against a list of possible names.
 * Includes a check to ensure possibleNames is an array.
 */
function findColumnIndex(headers, possibleNames) {
    if (!Array.isArray(possibleNames)) {
        console.error("[DEBUG] findColumnIndex received non-array for possibleNames:", possibleNames);
        return -1;
    }
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

/**
 * Finds the row of the current selection, and toggles a yellow highlight on the cells
 * in that row between the "StudentName" and "Outreach" columns.
 */
async function toggleHighlight(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("rowIndex");
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowIndex", "values"]);
      
      await context.sync();

      const headers = usedRange.values[0];
      const lowerCaseHeaders = headers.map(header => String(header || '').toLowerCase());
      const studentNameColIndex = findColumnIndex(lowerCaseHeaders, CONSTANTS.STUDENT_NAME_COLS);
      const outreachColIndex = findColumnIndex(lowerCaseHeaders, CONSTANTS.OUTREACH_COLS);

      if (studentNameColIndex === -1 || outreachColIndex === -1) {
        console.error("Could not find 'StudentName' and/or 'Outreach' columns.");
        return; 
      }
      
      const startCol = Math.min(studentNameColIndex, outreachColIndex);
      const endCol = Math.max(studentNameColIndex, outreachColIndex);
      const colCount = endCol - startCol + 1;
      const targetRowIndex = selectedRange.rowIndex;

      if (targetRowIndex < usedRange.rowIndex) return;

      const highlightRange = sheet.getRangeByIndexes(targetRowIndex, startCol, 1, colCount);
      highlightRange.format.fill.load("color");
      
      await context.sync();

      if (highlightRange.format.fill.color === "#FFFF00") {
        highlightRange.format.fill.clear();
      } else {
        highlightRange.format.fill.color = "yellow";
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error in toggleHighlight: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  } finally {
    if (event) {
      event.completed();
    }
  }
}

/**
 * Opens a dialog to transfer data to the clipboard.
 */
async function transferData(event) {
    let jsonDataString = "";
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            // Load both values and formulas to inspect hyperlinks
            usedRange.load("values, formulas");
            await context.sync();

            const headers = usedRange.values[0].map(header => String(header || '').toLowerCase());
            const colIndices = {
                studentName: findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS),
                gradeBook: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.gradeBook),
                daysOut: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.daysOut),
                lastLda: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.lastLda),
                grade: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.grade)
            };

            const dataToCopy = [];
            const hyperlinkRegex = /=HYPERLINK\("([^"]+)"/i;

            for (let i = 1; i < usedRange.values.length; i++) {
                const rowValues = usedRange.values[i];
                const rowFormulas = usedRange.formulas[i];
                const rowData = {};
                let hasData = false;

                if (colIndices.studentName !== -1 && rowValues[colIndices.studentName]) {
                    rowData.StudentName = rowValues[colIndices.studentName];
                    hasData = true;
                }

                if (colIndices.gradeBook !== -1 && rowValues[colIndices.gradeBook]) {
                    const formula = rowFormulas[colIndices.gradeBook];
                    const match = String(formula).match(hyperlinkRegex);
                    if (match && match[1]) {
                        rowData.GradeBook = match[1]; // Extract URL from formula
                    } else {
                        rowData.GradeBook = rowValues[colIndices.gradeBook]; // Fallback to value
                    }
                    hasData = true;
                }

                if (colIndices.daysOut !== -1 && rowValues[colIndices.daysOut]) {
                    rowData.DaysOut = rowValues[colIndices.daysOut];
                    hasData = true;
                }
                if (colIndices.lastLda !== -1 && rowValues[colIndices.lastLda]) {
                    rowData.LDA = rowValues[colIndices.lastLda];
                    hasData = true;
                }
                if (colIndices.grade !== -1 && rowValues[colIndices.grade]) {
                    rowData.Grade = rowValues[colIndices.grade];
                    hasData = true;
                }

                if (hasData) {
                    dataToCopy.push(rowData);
                }
            }

            if (dataToCopy.length > 0) {
                jsonDataString = JSON.stringify(dataToCopy, null, 2);
            }
        });

        if (!jsonDataString) {
            console.log("No data found to copy.");
            event.completed();
            return;
        }

        Office.context.ui.displayDialogAsync(
            'https://vsblanco.github.io/Student-Retention-Add-in/commands/transfer-dialog.html',
            { height: 60, width: 40, displayInIframe: true },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Transfer dialog failed to open: " + asyncResult.error.message);
                    event.completed();
                    return;
                }
                transferDialog = asyncResult.value;
                transferDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    const message = JSON.parse(arg.message);
                    if (message.type === 'dialogReady') {
                        transferDialog.messageChild(JSON.stringify({
                            type: 'dataForTransfer',
                            data: jsonDataString
                        }));
                    } else if (message.type === 'closeDialog') {
                        transferDialog.close();
                        transferDialog = null;
                    }
                });
                event.completed();
            }
        );
    } catch (error) {
        console.error("Error in transferData: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
        event.completed();
    }
}

/**
 * Opens a dialog to confirm creating a new LDA sheet.
 * @param {Office.AddinCommands.Event} event
 */
function openCreateLdaDialog(event) {
    Office.context.ui.displayDialogAsync(
        'https://vsblanco.github.io/Student-Retention-Add-in/commands/create-lda-dialog.html',
        { height: 25, width: 30, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Create LDA dialog failed to open: " + asyncResult.error.message);
                event.completed();
                return;
            }
            createLdaDialog = asyncResult.value;
            createLdaDialog.addEventHandler(Office.EventType.DialogMessageReceived, processCreateLdaMessage);
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
    console.log("[DEBUG] Starting handleCreateLdaSheet v12");
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
            masterRange.load("values, formulas");
            await context.sync();

            const masterData = masterRange.values;
            const masterFormulas = masterRange.formulas;
            const originalHeaders = masterData[0];
            const originalLCHeaders = originalHeaders.map(h => String(h || '').toLowerCase());
            
            const daysOutColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);
            if (daysOutColIdx === -1) throw new Error("'Days Out' column not found in Master List.");
            
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
                originalHeaders
            }
             // Autofit after hidng columns to avoid turning visible
            newSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            await context.sync();
            );

            if (includeFailingList) {
                console.log("[DEBUG] includeFailingList is true, creating failing list.");
                const nextStartRow = ldaTableEndRow > 0 ? ldaTableEndRow + 2 : 3;
                await createFailingListTable(context, newSheet, sheetName, nextStartRow, dataRowsWithIndex, masterFormulas, ldaColumns, hideLeftoverColumns, originalHeaders);
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

async function createFailingListTable(context, sheet, sheetName, startRow, masterDataWithIndex, masterFormulas, ldaColumns, hideLeftoverColumns, originalHeaders) {
    console.log("[DEBUG] Creating failing list table.");
    const originalLCHeaders = originalHeaders.map(h => String(h || '').toLowerCase());
    const gradeColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);

    if (gradeColIdx === -1) {
        console.warn("'Grade' column not found, cannot create failing list.");
        return;
    }

    const failingRows = masterDataWithIndex.filter(({ row }) => {
        const grade = row[gradeColIdx];
        return typeof grade === 'number' && (grade < 0.60 || (grade >= 1 && grade < 60));
    });
    failingRows.sort((a, b) => (a.row[gradeColIdx] || 0) - (b.row[gradeColIdx] || 0));

    if (failingRows.length > 0) {
        const titleRange = sheet.getRangeByIndexes(startRow, 0, 1, 1);
        titleRange.values = [["Failing Students"]];
        titleRange.format.font.bold = true;
        
        await createAndFormatTable(context, {
            sheet,
            tableName: sheetName.replace(/[^a-zA-Z0-9]/g, "_") + "_Failing",
            startRow: startRow + 1,
            dataRows: failingRows,
            masterFormulas,
            ldaColumns,
            hideLeftoverColumns,
            originalHeaders
        });
    }
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
    originalHeaders
  } = options;

  const finalHeaders = hideLeftoverColumns ? originalHeaders : ldaColumns;
  const indicesToKeep = finalHeaders.map(h => originalHeaders.indexOf(h));
  const originalLCHeaders = originalHeaders.map(h => String(h ?? '').toLowerCase());
  const gradeBookColIdx = findColumnIndex(originalLCHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);

  const dataToWrite = [];
  const formulasToWrite = [];

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

    // Load column names before accessing them
    table.columns.load("items/name");
    await context.sync();

    // Then hide columns
    if (hideLeftoverColumns) {
  console.log("[DEBUG] Hiding unused columns for table:", tableName);
  const selectedColumnsSet = new Set(ldaColumns.map(h => h.toLowerCase()));

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
        table.columns.getItemAt(colIdx).getRange().numberFormat = [["m/d/yyyy"]];
      }
    });

    return startRow + finalValues.length;
  } else {
    sheet.getRangeByIndexes(startRow, 0, 1, finalHeaders.length).values = [finalHeaders];
    return startRow + 1;
  }
}


// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("openImportDialog", openImportDialog);
Office.actions.associate("transferData", transferData);
Office.actions.associate("openCreateLdaDialog", openCreateLdaDialog);
//Version copiolotttttt
