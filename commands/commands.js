/*
 * This file contains the logic for ribbon button commands.
 * It is loaded by commands.html, which is specified as the FunctionFile in the manifest.
 */

const CONSTANTS = {
    // NOTE: "Student ID" and "Student Number" are treated as distinct values.
    // "Student ID" refers to the Canvas ID, used for creating gradebook links.
    // "Student Number" (and "Student Identifier") refers to the internal school ID.
    STUDENT_NAME_COLS: ["studentname", "student name"],
    OUTREACH_COLS: ["outreach"],
    STUDENT_ID_COLS: ["student id"],
    STUDENT_NUMBER_COLS: ["studentnumber", "student identifier"],
    MASTER_LIST_SHEET: "Master List",
    TEMPLATE_URL: 'https://vsblanco.github.io/Student-Retention-Add-in/Template.xlsx',
    COLUMN_MAPPINGS: {
        course: ["course"],
        courseId: ["course id"],
        currentScore: ["current score", "grade", "course grade"],
        grade: ["grade", "course grade"],
        gradeBook: ["grade book", "gradebook"],
        courseMissingAssignments: ["course missing assignments"],
        courseZeroAssignments: ["course zero assignments"]
    }
};

let importDialog = null;

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
        { height: 35, width: 35, displayInIframe: true },
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
    let hasCourseIdCol = false;
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
        hasCourseIdCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.courseId) !== -1;
        hasStudentNameCol = findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS) !== -1;
        hasCurrentScoreCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.currentScore) !== -1;
        console.log(`[DEBUG] Column checks: hasStudentIdCol=${hasStudentIdCol}, hasStudentNumberCol=${hasStudentNumberCol}, hasCourseIdCol=${hasCourseIdCol}, hasStudentNameCol=${hasStudentNameCol}, hasCurrentScoreCol=${hasCurrentScoreCol}`);

        if (hasStudentIdCol || hasStudentNameCol || hasStudentNumberCol) {
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
        
        const isGradeFile = hasStudentIdCol && hasCourseIdCol && hasStudentNameCol && hasCurrentScoreCol;
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
 * Handles the Master List update action. It updates rows in-place based on student name.
 * @param {object} message The message from the dialog.
 */
async function handleUpdateMaster(message) {
    if (importDialog) {
        importDialog.close();
    }
    console.log("Starting Master List update process...");
    try {
        // 1. Parse user's uploaded file
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
        console.log(`Parsed ${userData.length} data rows from imported file.`);

        const userStudentNameCol = findColumnIndex(lowerCaseUserHeaders, CONSTANTS.STUDENT_NAME_COLS);
        if (userStudentNameCol === -1) {
            throw new Error("Imported file is missing a 'Student Name' column.");
        }

        // 2. Update the "Master List" sheet
        await Excel.run(async (context) => {
            console.log("Accessing 'Master List' sheet...");
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount");
            await context.sync();
            console.log("'Master List' sheet loaded.");

            const masterHeaders = usedRange.values[0].map(h => String(h || ''));
            const lowerCaseMasterHeaders = masterHeaders.map(h => h.toLowerCase());
            const masterStudentNameCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            
            if (masterStudentNameCol === -1) {
                throw new Error("'Master List' is missing a 'StudentName' column.");
            }
            console.log("Found 'StudentName' column in 'Master List'.");

            // Create a map of student names in the Master List to their row index
            const masterNameMap = new Map();
            for (let i = 1; i < usedRange.values.length; i++) {
                const name = usedRange.values[i][masterStudentNameCol];
                if (name) {
                    masterNameMap.set(normalizeName(name), i);
                }
            }

            // Create a mapping from user file column index to master list column index
            const colMapping = lowerCaseUserHeaders.map(userHeader => 
                lowerCaseMasterHeaders.indexOf(userHeader)
            );

            let updatedCount = 0;
            const matchedStudentsLog = [];

            // Iterate through imported data and update matching rows
            for (const userRow of userData) {
                const studentName = userRow[userStudentNameCol];
                const normalizedName = normalizeName(studentName);

                if (masterNameMap.has(normalizedName)) {
                    updatedCount++;
                    const masterRowIndex = masterNameMap.get(normalizedName);
                    
                    // Update only the columns that exist in the imported file
                    for (let userColIdx = 0; userColIdx < userRow.length; userColIdx++) {
                        const masterColIdx = colMapping[userColIdx];
                        if (masterColIdx !== -1) {
                            const cell = sheet.getCell(masterRowIndex, masterColIdx);
                            cell.values = [[userRow[userColIdx] || ""]];
                        }
                    }

                    matchedStudentsLog.push({ student: studentName });
                }
            }
            
            console.log(`Found and updated ${updatedCount} matching students.`);
            console.table(matchedStudentsLog);

            if (updatedCount > 0) {
                sheet.getUsedRange().format.autofitColumns();
            }
            
            await context.sync();
            console.log("Master List update process completed successfully.");
        });

    } catch (error) {
        console.error("Error updating Master List: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Handles updating grades and gradebook links in the Master List.
 * If a student from the import is not on the Master List, they are added.
 * @param {object} message The message from the dialog.
 */
async function handleUpdateGrades(message) {
    if (importDialog) {
        importDialog.close();
    }
    console.log("Starting grade update process...");
    try {
        // 1. Parse user's uploaded file
        console.log("Parsing uploaded file for grade update...");
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
        console.log(`Parsed ${userData.length} rows from the imported file.`);

        // 2. Find column indices in user's file
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
        console.log("Found required columns for grade update in the imported file.");

        // 3. Create a map of student data from the imported file, keyed by normalized name, filtering out CAPV courses
        const studentDataMap = new Map();
        userData.forEach(row => {
            const courseName = row[userCourseCol] ? String(row[userCourseCol]) : '';
            if (courseName.toUpperCase().includes('CAPV')) {
                console.log(`[DEBUG] Skipping row for CAPV course: ${courseName}`);
                return; // Skip this row
            }

            const studentName = row[userStudentNameCol];
            if (studentName) {
                const normalized = normalizeName(studentName);
                studentDataMap.set(normalized, {
                    grade: row[userGradeCol],
                    courseId: row[userCourseIdCol],
                    studentId: row[userStudentIdCol],
                    originalName: studentName,
                    missingAssignments: row[userMissingAssignmentsCol],
                    zeroAssignments: row[userZeroAssignmentsCol]
                });
            }
        });
        console.log(`Created a map of ${studentDataMap.size} students from the imported file after filtering.`);

        // 4. Update the "Master List" sheet
        await Excel.run(async (context) => {
            console.log("Accessing 'Master List' sheet...");
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount");
            await context.sync();
            console.log("'Master List' sheet loaded.");

            const masterHeaders = usedRange.values[0].map(h => String(h || '').toLowerCase());
            const masterStudentNameCol = findColumnIndex(masterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterGradeCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);
            const masterGradebookCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            const masterMissingAssignmentsCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const masterZeroAssignmentsCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseZeroAssignments);
            
            if (masterStudentNameCol === -1 || masterGradeCol === -1 || masterGradebookCol === -1) {
                throw new Error("'Master List' is missing required columns: StudentName, Grade, or Grade Book.");
            }
            console.log("Found required columns in 'Master List'.");

            // Create a map of existing student names in the Master List to their row index
            const masterNameMap = new Map();
            for (let i = 1; i < usedRange.values.length; i++) {
                const name = usedRange.values[i][masterStudentNameCol];
                if (name) {
                    masterNameMap.set(normalizeName(name), i);
                }
            }

            let updatedCount = 0;
            let addedCount = 0;
            const newStudents = [];

            // Separate students into update list and add list
            for (const [normalizedName, importedData] of studentDataMap.entries()) {
                if (masterNameMap.has(normalizedName)) {
                    const masterRowIndex = masterNameMap.get(normalizedName);
                    
                    const gradeCell = sheet.getCell(masterRowIndex, masterGradeCol);
                    gradeCell.values = [[importedData.grade]];

                    if (importedData.courseId && importedData.studentId) {
                        const gradebookCell = sheet.getCell(masterRowIndex, masterGradebookCol);
                        const newGradebookLink = `https://nuc.instructure.com/courses/${importedData.courseId}/grades/${importedData.studentId}`;
                        const hyperlinkFormula = `=HYPERLINK("${newGradebookLink}", "Gradebook")`;
                        gradebookCell.formulas = [[hyperlinkFormula]];
                    }

                    if (masterMissingAssignmentsCol !== -1 && importedData.missingAssignments !== undefined) {
                        const missingCell = sheet.getCell(masterRowIndex, masterMissingAssignmentsCol);
                        missingCell.values = [[importedData.missingAssignments]];
                    }
                    if (masterZeroAssignmentsCol !== -1 && importedData.zeroAssignments !== undefined) {
                        const zeroCell = sheet.getCell(masterRowIndex, masterZeroAssignmentsCol);
                        zeroCell.values = [[importedData.zeroAssignments]];
                    }
                    
                    updatedCount++;
                } else {
                    newStudents.push(importedData);
                }
            }

            // Add new students to the end of the sheet
            if (newStudents.length > 0) {
                const lastRow = usedRange.rowCount;
                for (let i = 0; i < newStudents.length; i++) {
                    const student = newStudents[i];
                    const newRowIndex = lastRow + i;
                    
                    const nameCell = sheet.getCell(newRowIndex, masterStudentNameCol);
                    nameCell.values = [[formatToLastFirst(student.originalName)]];

                    const gradeCell = sheet.getCell(newRowIndex, masterGradeCol);
                    gradeCell.values = [[student.grade]];

                    if (student.courseId && student.studentId) {
                        const newGradebookLink = `https://nuc.instructure.com/courses/${student.courseId}/grades/${student.studentId}`;
                        const hyperlinkFormula = `=HYPERLINK("${newGradebookLink}", "Gradebook")`;
                        const gradebookCell = sheet.getCell(newRowIndex, masterGradebookCol);
                        gradebookCell.formulas = [[hyperlinkFormula]];
                    }

                    if (masterMissingAssignmentsCol !== -1 && student.missingAssignments !== undefined) {
                        const missingCell = sheet.getCell(newRowIndex, masterMissingAssignmentsCol);
                        missingCell.values = [[student.missingAssignments]];
                    }
                    if (masterZeroAssignmentsCol !== -1 && student.zeroAssignments !== undefined) {
                        const zeroCell = sheet.getCell(newRowIndex, masterZeroAssignmentsCol);
                        zeroCell.values = [[student.zeroAssignments]];
                    }

                    // Highlight the new row
                    const newRowRange = sheet.getRangeByIndexes(newRowIndex, 0, 1, masterHeaders.length);
                    newRowRange.format.fill.color = "#ADD8E6"; // Light Blue
                    addedCount++;
                }
            }
            
            console.log(`Found and updated ${updatedCount} matching students.`);
            console.log(`Added ${addedCount} new students.`);

            if (updatedCount > 0 || addedCount > 0) {
                sheet.getUsedRange().format.autofitColumns();
            }
            
            await context.sync();
            console.log("Grade update process completed successfully.");
        });

    } catch (error) {
        console.error("Error updating grades in Master List: " + error);
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
 */
function findColumnIndex(headers, possibleNames) {
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

// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("openImportDialog", openImportDialog);
