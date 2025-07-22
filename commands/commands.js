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
        { height: 30, width: 35, displayInIframe: true },
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
 * Checks if the file has a student ID and if the Master List sheet exists.
 * @param {object} message The message from the dialog.
 */
async function handleFileSelected(message) {
    const { fileName, data: dataUrl } = message;
    let hasStudentIdCol = false;
    let hasMasterListSheet = false;
    let hasCourseIdCol = false;

    try {
        const arrayBuffer = dataUrlToArrayBuffer(dataUrl);
        const workbook = new ExcelJS.Workbook();
        
        if (fileName.toLowerCase().endsWith('.xlsx')) {
            await workbook.xlsx.load(arrayBuffer);
        } else {
            const csvData = new TextDecoder("utf-8").decode(arrayBuffer);
            const rows = csvData.split(/\r?\n/).filter(row => row.trim().length > 0);
            const data = rows.map(row => parseCsvRow(row));
            // Manually create worksheet from CSV data for ExcelJS
            const worksheet = workbook.addWorksheet('sheet1');
            worksheet.addRows(data);
        }

        const worksheet = workbook.worksheets[0];
        const headers = (worksheet.getRow(1).values || []).map(h => String(h || '').toLowerCase());
        if (findColumnIndex(headers, CONSTANTS.STUDENT_ID_COLS) !== -1) {
            hasStudentIdCol = true;
        }
        if (findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.courseId) !== -1) {
            hasCourseIdCol = true;
        }

        if (hasStudentIdCol) {
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
                canUpdateGrades: hasStudentIdCol && hasMasterListSheet && hasCourseIdCol
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

        if (!templateHeaders || templateHeaders.length === 0) {
            throw new Error("Could not determine template headers from Template.xlsx or Master List.");
        }

        // 2. Parse user's uploaded file
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
        const userHeaders = (userWorksheet.getRow(1).values || []).slice(1).map(h => String(h || ''));
        const userData = [];
        userWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                userData.push((row.values || []).slice(1));
            }
        });

        // 3. Create column mapping
        const lowerCaseUserHeaders = userHeaders.map(h => h.toLowerCase());
        const colMapping = templateHeaders.map(templateHeader => {
            if (!templateHeader) return -1;
            return lowerCaseUserHeaders.indexOf(templateHeader.toLowerCase());
        });

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

        // 5. Write to "Master List" sheet, preserving the header
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            
            // Clear only the data, not the header
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();

            if (usedRange.rowCount > 1) {
                const dataRange = usedRange.getOffsetRange(1, 0).getResizedRange(usedRange.rowCount - 1, 0);
                dataRange.clear();
            }

            // Write the new data starting from the second row (index 1)
            if (finalData.length > 0) {
                const targetRange = sheet.getRangeByIndexes(1, 0, finalData.length, templateHeaders.length);
                targetRange.values = finalData;
                targetRange.format.autofitColumns();
            }
            
            await context.sync();
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
        const userHeaders = (userWorksheet.getRow(1).values || []).slice(1).map(h => String(h || '').toLowerCase());
        const userData = [];
        userWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                userData.push((row.values || []).slice(1));
            }
        });

        // 2. Find column indices in user's file
        const userStudentIdCol = findColumnIndex(userHeaders, CONSTANTS.STUDENT_ID_COLS);
        const userCourseIdCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.courseId);
        const userGradeCol = findColumnIndex(userHeaders, CONSTANTS.COLUMN_MAPPINGS.currentScore);

        if (userStudentIdCol === -1 || userCourseIdCol === -1 || userGradeCol === -1) {
            throw new Error("Imported file is missing one of the required columns: Student ID, Course ID, or Current Score/Grade.");
        }

        // 3. Create a map of student data from the imported file
        const studentDataMap = new Map();
        userData.forEach(row => {
            const studentId = row[userStudentIdCol];
            if (studentId) {
                studentDataMap.set(String(studentId), {
                    grade: row[userGradeCol],
                    courseId: row[userCourseIdCol]
                });
            }
        });

        // 4. Update the "Master List" sheet
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount, columnCount");
            await context.sync();

            const masterHeaders = usedRange.values[0].map(h => String(h || '').toLowerCase());
            const masterStudentIdCol = findColumnIndex(masterHeaders, CONSTANTS.STUDENT_ID_COLS);
            const masterGradeCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);
            const masterGradebookCol = findColumnIndex(masterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            
            if (masterStudentIdCol === -1 || masterGradeCol === -1 || masterGradebookCol === -1) {
                throw new Error("'Master List' is missing required columns: Student ID, Grade, or Grade Book.");
            }

            const masterData = usedRange.values;
            const updatedData = [];

            // Start from 1 to skip header
            for (let i = 1; i < masterData.length; i++) {
                const row = masterData[i];
                const studentId = String(row[masterStudentIdCol]);
                if (studentDataMap.has(studentId)) {
                    const importedData = studentDataMap.get(studentId);
                    row[masterGradeCol] = importedData.grade;
                    row[masterGradebookCol] = `https://nuc.instructure.com/courses/${importedData.courseId}/grades/${studentId}`;
                }
                updatedData.push(row);
            }
            
            if (updatedData.length > 0) {
                const dataRange = sheet.getRangeByIndexes(1, 0, updatedData.length, masterHeaders.length);
                dataRange.values = updatedData;
                dataRange.format.autofitColumns();
            }
            
            await context.sync();
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
