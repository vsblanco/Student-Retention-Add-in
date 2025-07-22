/*
 * This file contains the logic for ribbon button commands.
 * It is loaded by commands.html, which is specified as the FunctionFile in the manifest.
 */

const CONSTANTS = {
    // Column Header Names
    STUDENT_NAME_COLS: ["studentname", "student name"],
    OUTREACH_COLS: ["outreach"],
    STUDENT_ID_COLS: ["student id", "studentnumber", "student identifier"],
    MASTER_LIST_SHEET: "Master List",
    TEMPLATE_URL: 'https://vsblanco.github.io/Student-Retention-Add-in/Template.xlsx',
    COLUMN_MAPPINGS: {
        courseId: ["course id"],
        currentScore: ["current score", "grade"],
        grade: ["grade", "course grade"],
        gradeBook: ["grade book", "gradebook"]
    }
};

let importDialog = null;

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
    let hasStudentIdCol = false;
    let hasMasterListSheet = false;
    let hasCourseIdCol = false;
    let hasStudentNameCol = false;

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

        const worksheet = workbook.worksheets[0];
        const headers = [];
        worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell) => {
            headers.push(String(cell.value || '').toLowerCase());
        });
        
        hasStudentIdCol = findColumnIndex(headers, CONSTANTS.STUDENT_ID_COLS) !== -1;
        hasCourseIdCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.courseId) !== -1;
        hasStudentNameCol = findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS) !== -1;

        if (hasStudentIdCol || hasStudentNameCol) {
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

        if (importDialog) {
            importDialog.messageChild(JSON.stringify({ 
                canUpdateMaster: hasStudentIdCol && hasMasterListSheet,
                canUpdateGrades: hasStudentIdCol && hasMasterListSheet && hasCourseIdCol && hasStudentNameCol
            }));
        }
    } catch (error) {
        console.error("Error during file check:", error);
    }
}


/**
 * Handles the Master List update action.
 * @param {object} message The message from the dialog.
 */
async function handleUpdateMaster(message) {
    if (importDialog) {
        importDialog.close();
    }
    console.log("Starting Master List update process...");
    try {
        let templateHeaders;

        // 1. Fetch and parse Template.xlsx, with a fallback to the Master List headers
        try {
            const templateResponse = await fetch(CONSTANTS.TEMPLATE_URL);
            if (!templateResponse.ok) throw new Error('Failed to fetch Template.xlsx');
            const templateArrayBuffer = await templateResponse.arrayBuffer();
            const templateWorkbook = new ExcelJS.Workbook();
            await templateWorkbook.xlsx.load(templateArrayBuffer);
            const templateWorksheet = templateWorkbook.worksheets[0];
            templateHeaders = (templateWorksheet.getRow(1).values || []).slice(1).map(h => String(h || ''));
            if (templateHeaders.length === 0) {
                throw new Error("Template.xlsx is empty or has no headers.");
            }
        } catch (templateError) {
            console.error("Could not load or parse Template.xlsx. Falling back to 'Master List' headers. Error:", templateError);
            
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
                const usedRange = sheet.getUsedRange(true);
                usedRange.load("values, rowCount");
                await context.sync();

                if (usedRange.rowCount > 0 && usedRange.values && usedRange.values[0].some(h => h)) {
                    templateHeaders = usedRange.values[0].map(h => String(h || ''));
                } else {
                    throw new Error("'Master List' sheet is empty and Template.xlsx could not be loaded. Please ensure one of them has headers.");
                }
            });
        }
        console.log("Using template headers:", templateHeaders);

        if (!templateHeaders || templateHeaders.length === 0) {
            throw new Error("Could not determine template headers from Template.xlsx or Master List.");
        }

        // 2. Parse user's uploaded file
        console.log("Parsing uploaded file for Master List update...");
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
        console.log("Parsed imported file headers:", userHeaders);

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

        // 3. Create column mapping
        const lowerCaseUserHeaders = userHeaders.map(h => h.toLowerCase());
        const colMapping = templateHeaders.map(templateHeader => {
            if (!templateHeader) return -1;
            return lowerCaseUserHeaders.indexOf(templateHeader.toLowerCase());
        });
        console.log("Created column mapping:", colMapping);

        // 4. Create final data array based on template (data only, no header)
        const finalData = [];
        userData.forEach(userRow => {
            const newRow = new Array(templateHeaders.length).fill("");
            colMapping.forEach((userColIndex, templateColIndex) => {
                if (userColIndex !== -1 && userColIndex < userRow.length) {
                    newRow[templateColIndex] = userRow[userColIndex] || "";
                }
            });
            finalData.push(newRow);
        });
        console.log(`Created ${finalData.length} rows of data to be written to Master List.`);

        // 5. Write to "Master List" sheet, preserving the header
        await Excel.run(async (context) => {
            console.log("Accessing 'Master List' to write data...");
            let sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();
            console.log(`'Master List' has ${usedRange.rowCount} rows before clearing.`);

            if (usedRange.rowCount > 1) {
                const dataRange = usedRange.getOffsetRange(1, 0).getResizedRange(usedRange.rowCount - 1, 0);
                dataRange.clear();
                console.log("Cleared existing data from 'Master List'.");
            }

            if (finalData.length > 0) {
                const targetRange = sheet.getRangeByIndexes(1, 0, finalData.length, templateHeaders.length);
                targetRange.values = finalData;
                console.log(`Writing ${finalData.length} new rows to 'Master List'.`);
                sheet.getUsedRange().format.autofitColumns();
            } else {
                console.log("No data to write.");
            }
            
            await context.sync();
            console.log("Master List update process completed.");
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
 * @param {object} message The message from the dialog.
 */
async function handleUpdateGrades(message) {
    if (importDialog) {
        importDialog.close();
    }
    console.log("Starting grade update process...");
    try {
        // Helper to normalize names from "Last, First" or "First Last" to "first last"
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
                row.eachCell({ includeEmpty: true }, (cell) => {
                    rowData.push(cell.value);
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

        if (userStudentIdCol === -1 || userCourseIdCol === -1 || userGradeCol === -1 || userStudentNameCol === -1) {
            throw new Error("Imported file is missing one of the required columns: Student Name, Student ID, Course ID, or Current Score/Grade.");
        }
        console.log("Found required columns in the imported file.");

        // 3. Create a map of student data from the imported file, keyed by normalized name
        const studentDataMap = new Map();
        userData.forEach(row => {
            const studentName = row[userStudentNameCol];
            if (studentName) {
                const normalized = normalizeName(studentName);
                studentDataMap.set(normalized, {
                    grade: row[userGradeCol],
                    courseId: row[userCourseIdCol],
                    studentId: row[userStudentIdCol]
                });
            }
        });
        console.log(`Created a map of ${studentDataMap.size} students from the imported file.`);

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
            
            if (masterStudentNameCol === -1 || masterGradeCol === -1 || masterGradebookCol === -1) {
                throw new Error("'Master List' is missing required columns: StudentName, Grade, or Grade Book.");
            }
            console.log("Found required columns in 'Master List'.");

            const masterData = usedRange.values;
            let updatedCount = 0;

            console.log(`Iterating through ${masterData.length - 1} rows in 'Master List' to find matches...`);
            // Start from 1 to skip header
            for (let i = 1; i < masterData.length; i++) {
                const row = masterData[i];
                const masterStudentName = row[masterStudentNameCol];
                const normalizedMasterName = normalizeName(masterStudentName);

                if (studentDataMap.has(normalizedMasterName)) {
                    updatedCount++;
                    const importedData = studentDataMap.get(normalizedMasterName);
                    
                    const gradeCell = sheet.getRangeByIndexes(i, masterGradeCol, 1, 1);
                    const gradebookCell = sheet.getRangeByIndexes(i, masterGradebookCol, 1, 1);
                    
                    gradeCell.values = [[importedData.grade]];
                    
                    const gradebookLink = `https://nuc.instructure.com/courses/${importedData.courseId}/grades/${importedData.studentId}`;
                    gradebookCell.values = [[gradebookLink]];
                    
                    console.log(`Updating row ${i + 1} for student: ${masterStudentName}`);
                }
            }
            
            console.log(`Found and updated ${updatedCount} matching students.`);
            if (updatedCount > 0) {
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
