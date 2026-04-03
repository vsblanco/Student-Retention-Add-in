/*
 * background-service.js
 *
 * Main entry point and background service for ribbon button commands.
 *
 * This file:
 * - Registers all ribbon button actions with Office.actions
 * - Runs continuously as a background script (even when taskpanes are closed)
 * - Manages Chrome Extension Service communication
 * - Handles Master List data synchronization
 */
import { toggleHighlight, transferData, createSendToCallQueue } from './ribbon-actions.js';
import chromeExtensionService from '../react/src/services/chromeExtensionService.js';
import { CONSTANTS, findColumnIndex, normalizeName, formatToLastFirst, parseDate } from './shared-utilities.js';

// Batch size for chunked read/write operations to avoid Excel's ~5MB payload limit.
// 500 rows is consistent with ldaProcessor.js and safely under the limit for ~37 columns.
const BATCH_SIZE = 500;
// Smaller batch size for formatting operations (cell colors).
const FORMAT_BATCH_SIZE = 50;

/**
 * Sends an import status update back to the Chrome extension via window.postMessage.
 * The extension's content script (excelConnector.js) picks this up and forwards it.
 * @param {string} status - Status key (e.g., 'received', 'writing_data', 'complete', 'error')
 * @param {string} message - Human-readable status message
 * @param {number} [progress] - Optional progress percentage (0-100)
 * @param {string} [error] - Optional error message (for error status)
 */
function sendImportStatus(status, message, progress, error) {
    try {
        window.postMessage({
            type: "SRK_IMPORT_STATUS",
            data: {
                status: status,
                message: message,
                ...(progress != null && { progress }),
                ...(error && { error })
            }
        }, "*");
    } catch (e) {
        console.warn("Failed to send import status:", e);
    }
}

// This function is required for the Analytics button, even if it does nothing,
// because the manifest uses a ShowTaskpane action.
function openAnalyticsPane(event) {
    event.completed();
}

// Called automatically when a new document is opened
// This ensures the background service starts on document load
function onDocumentOpen(event) {
    console.log("Document opened - background services are active");
    event.completed();
}

/**
 * Imports missing assignments data to the Missing Assignments sheet
 * @param {Array} studentsWithAssignments - Array of student objects that have missingAssignments
 * @returns {Promise<void>}
 */
async function importMissingAssignments(studentsWithAssignments) {
    try {
        console.log("ImportMissingAssignments: Starting import of missing assignments...");

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Check if Missing Assignments sheet exists, create if not
            let maSheet = sheets.items.find(s => s.name === "Missing Assignments");
            if (!maSheet) {
                console.log("ImportMissingAssignments: Creating Missing Assignments sheet...");
                maSheet = sheets.add("Missing Assignments");
                await context.sync();
            } else {
                maSheet = context.workbook.worksheets.getItem("Missing Assignments");
            }

            // Define the column structure
            const headers = ["Student", "Grade", "Grade Book", "Assignment", "Due Date", "Score", "Submission"];

            // Clear existing content
            const usedRange = maSheet.getUsedRangeOrNullObject();
            usedRange.load("address");
            await context.sync();

            if (!usedRange.isNullObject) {
                usedRange.clear(Excel.ClearApplyTo.all);
            }

            // Write headers
            const headerRange = maSheet.getRangeByIndexes(0, 0, 1, headers.length);
            headerRange.values = [headers];
            headerRange.format.font.bold = true;
            await context.sync();

            // Prepare data rows
            const dataToWrite = [];
            const formulasToWrite = [];

            studentsWithAssignments.forEach(student => {
                const studentName = student["Student Name"] || student.StudentName || "";
                const grade = student.Grade ?? "";
                const gradeBookUrl = typeof student["Grade Book"] === 'object' ? student["Grade Book"].url : student["Grade Book"];
                const gradeBookText = typeof student["Grade Book"] === 'object' ? student["Grade Book"].text : "Grade Book";

                if (student.missingAssignments && Array.isArray(student.missingAssignments)) {
                    student.missingAssignments.forEach(assignment => {
                        const row = new Array(headers.length).fill("");
                        const formulaRow = new Array(headers.length).fill(null);

                        // Student
                        row[0] = studentName;

                        // Grade
                        row[1] = grade;

                        // Grade Book (HYPERLINK)
                        if (gradeBookUrl) {
                            formulaRow[2] = `=HYPERLINK("${gradeBookUrl}", "${gradeBookText}")`;
                            row[2] = gradeBookText;
                        }

                        // Assignment (HYPERLINK with assignment URL and title as friendly name)
                        const assignmentUrl = typeof assignment.Assignment === 'object' ? assignment.Assignment.url : assignment.assignmentLink || assignment.assignmentUrl;
                        const assignmentTitle = typeof assignment.Assignment === 'object' ? assignment.Assignment.text : assignment.assignmentTitle || assignment.assignmentName || "Assignment";
                        if (assignmentUrl) {
                            formulaRow[3] = `=HYPERLINK("${assignmentUrl}", "${assignmentTitle}")`;
                            row[3] = assignmentTitle;
                        } else {
                            row[3] = assignmentTitle;
                        }

                        // Due Date
                        row[4] = typeof assignment["Due Date"] === 'object' ? assignment["Due Date"].text : assignment.dueDate || assignment["Due Date"] || "";

                        // Score
                        row[5] = typeof assignment.Score === 'object' ? assignment.Score.text : assignment.score ?? assignment.Score ?? "";

                        // Submission (HYPERLINK with submission URL and "Missing" as friendly name)
                        const submissionUrl = typeof assignment.Submission === 'object' ? assignment.Submission.url : assignment.submissionLink || assignment.submissionUrl;
                        if (submissionUrl) {
                            formulaRow[6] = `=HYPERLINK("${submissionUrl}", "Missing")`;
                            row[6] = "Missing";
                        }

                        dataToWrite.push(row);
                        formulasToWrite.push(formulaRow);
                    });
                }
            });

            if (dataToWrite.length > 0) {
                const totalBatches = Math.ceil(dataToWrite.length / BATCH_SIZE);
                console.log(`ImportMissingAssignments: Writing ${dataToWrite.length} assignment rows in ${totalBatches} batch(es)...`);

                for (let batchStart = 0; batchStart < dataToWrite.length; batchStart += BATCH_SIZE) {
                    const batchEnd = Math.min(batchStart + BATCH_SIZE, dataToWrite.length);
                    const batchSize = batchEnd - batchStart;

                    const batchValues = dataToWrite.slice(batchStart, batchEnd);
                    const batchFormulas = formulasToWrite.slice(batchStart, batchEnd);

                    const batchRange = maSheet.getRangeByIndexes(1 + batchStart, 0, batchSize, headers.length);
                    batchRange.values = batchValues;
                    batchRange.formulas = batchFormulas;

                    await context.sync();

                    const currentBatch = Math.floor(batchStart / BATCH_SIZE) + 1;
                    console.log(`ImportMissingAssignments: Batch ${currentBatch}/${totalBatches} completed (rows ${batchStart + 1}-${batchEnd})`);
                }

                // Autofit columns after all data is written
                maSheet.getUsedRange().format.autofitColumns();
                await context.sync();

                console.log(`ImportMissingAssignments: Successfully imported ${dataToWrite.length} missing assignments`);
            } else {
                console.log("ImportMissingAssignments: No missing assignments to import");
            }
        });

    } catch (error) {
        console.error("ImportMissingAssignments: Error importing missing assignments:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("ImportMissingAssignments: Debug info:", JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Shows a dialog prompting the user to create a Master List sheet when one doesn't exist.
 * If the user confirms, creates the sheet and retries the import.
 * @param {object} payload - The original import payload to retry after sheet creation.
 */
function showMissingMasterListDialog(payload) {
    const dialogUrl = 'https://vsblanco.github.io/Student-Retention-Add-in/background-services/missing-masterlist-dialog.html';

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 40, width: 35, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("MissingMasterList: Dialog failed to open:", asyncResult.error.message);
                return;
            }
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                const message = JSON.parse(arg.message);
                dialog.close();
                if (message.type === 'createMasterList') {
                    console.log("MissingMasterList: User chose to create Master List sheet.");
                    await createMasterListAndImport(payload);
                } else {
                    console.log("MissingMasterList: User cancelled. Import aborted.");
                }
            });
        }
    );
}

/**
 * Creates a blank Master List sheet and then retries the import.
 * @param {object} payload - The original import payload.
 */
async function createMasterListAndImport(payload) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.add(CONSTANTS.MASTER_LIST_SHEET);
            // Write the incoming headers as the first row so the import has something to work with
            if (payload.headers && payload.headers.length > 0) {
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, payload.headers.length);
                headerRange.values = [payload.headers];
            }
            sheet.activate();
            await context.sync();
            console.log("MissingMasterList: Master List sheet created successfully.");
        });

        // Retry the import now that the sheet exists
        await importMasterListFromExtension(payload);
    } catch (error) {
        console.error("MissingMasterList: Error creating Master List sheet:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("MissingMasterList: Debug info:", JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Imports master list data received from the Chrome extension
 * @param {object} payload - The data payload from the extension
 * @param {string[]} payload.headers - Array of column headers
 * @param {Array[]} payload.data - Array of data rows
 * @returns {Promise<void>}
 */
async function importMasterListFromExtension(payload) {
    try {
        console.log("ImportFromExtension: Starting master list import from Chrome extension...");

        // Validate payload
        if (!payload || !payload.headers || !payload.data) {
            console.error("ImportFromExtension: Invalid payload - missing headers or data");
            return;
        }

        const { headers: incomingHeaders, data: incomingData } = payload;
        console.log(`ImportFromExtension: Received ${incomingData.length} rows with headers: [${incomingHeaders.join(', ')}]`);

        sendImportStatus('received', `Received ${incomingData.length} students, validating data...`);

        await Excel.run(async (context) => {
            // Check if Master List sheet exists
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const masterListSheet = sheets.items.find(s => s.name === CONSTANTS.MASTER_LIST_SHEET);
            if (!masterListSheet) {
                console.log("ImportFromExtension: Master List sheet not found. Prompting user to create one.");
                showMissingMasterListDialog(payload);
                return;
            }

            console.log("ImportFromExtension: Master List sheet found, proceeding with import...");
            sendImportStatus('reading_data', 'Reading existing spreadsheet data...');

            // Get the Master List sheet
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            // Load only dimensions first to avoid payload size limit
            usedRange.load("rowCount, columnCount, rowIndex, columnIndex");
            await context.sync();

            const totalRows = usedRange.rowCount;
            const totalCols = usedRange.columnCount;

            // Read values and formulas in batches to avoid response payload limits
            let allValues = [];
            let allFormulas = [];

            for (let startRow = 0; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = sheet.getRangeByIndexes(usedRange.rowIndex + startRow, usedRange.columnIndex, rowsToRead, totalCols);
                batchRange.load("values, formulas");
                await context.sync();

                allValues = allValues.concat(batchRange.values);
                allFormulas = allFormulas.concat(batchRange.formulas);

                if (totalRows > BATCH_SIZE) {
                    const currentBatch = Math.floor(startRow / BATCH_SIZE) + 1;
                    const totalBatchCount = Math.ceil(totalRows / BATCH_SIZE);
                    console.log(`ImportFromExtension: Read batch ${currentBatch}/${totalBatchCount} (rows ${startRow + 1}-${startRow + rowsToRead})`);
                }
            }

            console.log(`ImportFromExtension: Read ${allValues.length} rows in ${Math.ceil(totalRows / BATCH_SIZE)} batch(es)`);

            // Get Master List headers
            const masterHeaders = allValues[0].map(h => String(h || ''));
            const lowerCaseMasterHeaders = masterHeaders.map(h => h.toLowerCase());
            const lowerCaseIncomingHeaders = incomingHeaders.map(h => String(h || '').toLowerCase());

            console.log(`ImportFromExtension: Master List headers: [${masterHeaders.join(', ')}]`);

            // Helper function to normalize headers (remove all whitespace)
            const normalizeHeader = (header) => String(header).toLowerCase().replace(/\s+/g, '');

            // Create normalized versions for matching
            const normalizedMasterHeaders = masterHeaders.map(h => normalizeHeader(h));
            const normalizedIncomingHeaders = incomingHeaders.map(h => normalizeHeader(h));

            // Create column mapping (incoming column index -> master column index)
            // First try exact case-insensitive match, then try normalized match
            const colMapping = lowerCaseIncomingHeaders.map((incomingHeader, idx) => {
                // Try exact match first
                let masterIdx = lowerCaseMasterHeaders.indexOf(incomingHeader);

                // If no exact match, try normalized match (removes whitespace differences)
                if (masterIdx === -1) {
                    masterIdx = normalizedMasterHeaders.indexOf(normalizedIncomingHeaders[idx]);
                }

                return masterIdx;
            });

            // Detect new columns from incoming data not yet on Master List
            const newColumns = [];
            for (let i = 0; i < incomingHeaders.length; i++) {
                if (colMapping[i] === -1) {
                    const newColIdx = masterHeaders.length;
                    masterHeaders.push(String(incomingHeaders[i] || ''));
                    lowerCaseMasterHeaders.push(lowerCaseIncomingHeaders[i]);
                    normalizedMasterHeaders.push(normalizedIncomingHeaders[i]);
                    colMapping[i] = newColIdx;
                    newColumns.push(String(incomingHeaders[i] || ''));
                }
            }

            if (newColumns.length > 0) {
                console.log(`ImportFromExtension: Adding ${newColumns.length} new column(s) to Master List: [${newColumns.join(', ')}]`);
            }

            console.log(`ImportFromExtension: Column mapping: [${colMapping.join(', ')}]`);

            // Determine which master columns are NOT covered by the incoming data
            const matchedMasterCols = new Set(colMapping);
            const unmatchedMasterCols = [];
            for (let i = 0; i < masterHeaders.length; i++) {
                if (!matchedMasterCols.has(i)) {
                    unmatchedMasterCols.push(i);
                }
            }

            // Find the student name column in both incoming data and master list
            const incomingStudentNameCol = findColumnIndex(lowerCaseIncomingHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterStudentNameCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterGradebookCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            const masterAssignedCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.assigned);
            const masterMissingAssignmentsCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const incomingMissingAssignmentsCol = findColumnIndex(lowerCaseIncomingHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const masterIdCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.STUDENT_ID_COLS);
            const incomingIdCol = findColumnIndex(lowerCaseIncomingHeaders, CONSTANTS.STUDENT_ID_COLS);

            if (incomingStudentNameCol === -1) {
                console.error("ImportFromExtension: Incoming data is missing a 'Student Name' column");
                return;
            }

            if (masterStudentNameCol === -1) {
                console.error("ImportFromExtension: Master List is missing a 'StudentName' column");
                return;
            }

            // Create a map of existing students to preserve certain data
            const masterDataMap = new Map();
            const valueToColorMap = new Map();

            // First pass: Get existing data to preserve
            for (let i = 1; i < allValues.length; i++) {
                const name = allValues[i][masterStudentNameCol];
                if (name) {
                    const normalizedName = normalizeName(name);
                    const gradebookFormula = (masterGradebookCol !== -1 && allFormulas[i][masterGradebookCol]) ? allFormulas[i][masterGradebookCol] : null;
                    const assignedValue = (masterAssignedCol !== -1) ? allValues[i][masterAssignedCol] : null;

                    masterDataMap.set(normalizedName, {
                        gradebookFormula: gradebookFormula,
                        assigned: assignedValue
                    });
                }
            }

            console.log(`ImportFromExtension: Created map of ${masterDataMap.size} existing students`);
            sendImportStatus('preserving_data', `Preserving data for ${masterDataMap.size} existing students...`);

            // Build preservation map for master columns not covered by the incoming data
            const preservationMap = new Map();
            const nonBlankUnmatchedCols = [];
            if (unmatchedMasterCols.length > 0) {
                // Check which unmatched columns have any non-blank data
                for (const colIdx of unmatchedMasterCols) {
                    let hasData = false;
                    for (let i = 1; i < allValues.length; i++) {
                        const val = allValues[i][colIdx];
                        if (val !== null && val !== undefined && String(val).trim() !== '') {
                            hasData = true;
                            break;
                        }
                    }
                    if (hasData) nonBlankUnmatchedCols.push(colIdx);
                }

                // Build preservation data per student (keyed by Student ID, fallback to name)
                if (nonBlankUnmatchedCols.length > 0) {
                    for (let i = 1; i < allValues.length; i++) {
                        let key;
                        if (masterIdCol !== -1) {
                            const idValue = allValues[i][masterIdCol];
                            if (idValue != null && String(idValue).trim() !== '') {
                                key = String(idValue).trim();
                            }
                        }
                        if (!key) {
                            const name = allValues[i][masterStudentNameCol];
                            if (name) key = normalizeName(name);
                        }
                        if (!key) continue;

                        const preservedValues = {};
                        const preservedFormulas = {};
                        for (const colIdx of nonBlankUnmatchedCols) {
                            const val = allValues[i][colIdx];
                            const formula = allFormulas[i][colIdx];
                            if (formula && typeof formula === 'string' && formula.startsWith('=')) {
                                preservedFormulas[colIdx] = formula;
                                preservedValues[colIdx] = val;
                            } else if (val !== null && val !== undefined && String(val).trim() !== '') {
                                preservedValues[colIdx] = val;
                            }
                        }
                        preservationMap.set(key, { values: preservedValues, formulas: preservedFormulas });
                    }
                    console.log(`ImportFromExtension: Built preservation map for ${nonBlankUnmatchedCols.length} unmatched column(s): [${nonBlankUnmatchedCols.map(i => masterHeaders[i]).join(', ')}]`);
                }
            }

            // Get colors for assigned column values
            if (masterAssignedCol !== -1) {
                const allAssignedValues = allValues.map(row => row[masterAssignedCol]);
                const uniqueValues = [...new Set(allAssignedValues.slice(1).filter(v => v && String(v).trim() !== ""))];

                if (uniqueValues.length > 0) {
                    console.log(`ImportFromExtension: Found ${uniqueValues.length} unique values in 'Assigned' column`);
                    const cellsToLoad = [];
                    uniqueValues.forEach(value => {
                        const firstInstanceIndex = allAssignedValues.indexOf(value);
                        if (firstInstanceIndex > 0) {
                            const absoluteRowIndex = usedRange.rowIndex + firstInstanceIndex;
                            const absoluteColIndex = usedRange.columnIndex + masterAssignedCol;
                            const cell = sheet.getCell(absoluteRowIndex, absoluteColIndex);
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
                    console.log(`ImportFromExtension: Cached colors for ${valueToColorMap.size} unique values`);
                }
            }

            // Categorize students as new or existing
            const newStudents = [];
            const existingStudents = [];

            for (const row of incomingData) {
                const studentName = row[incomingStudentNameCol];
                const normalizedName = normalizeName(studentName);

                if (masterDataMap.has(normalizedName)) {
                    existingStudents.push(row);
                } else {
                    newStudents.push(row);
                }
            }

            console.log(`ImportFromExtension: Found ${newStudents.length} new students and ${existingStudents.length} existing students`);

            // Clear the sheet (keeping headers)
            if (totalRows > 1) {
                const rangeToClear = sheet.getRangeByIndexes(1, 0, totalRows - 1, masterHeaders.length);
                rangeToClear.clear(Excel.ClearApplyTo.all);
                rangeToClear.getEntireRow().delete(Excel.DeleteShiftDirection.up);
                await context.sync();
                console.log("ImportFromExtension: Sheet cleared");
            }

            // Update header row to include any new columns from import
            if (newColumns.length > 0) {
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, masterHeaders.length);
                headerRange.values = [masterHeaders];
                headerRange.format.font.bold = true;
                await context.sync();
                console.log(`ImportFromExtension: Updated header row with ${newColumns.length} new column(s)`);
            }

            // Combine students (new first, then existing)
            const allStudentsToWrite = [...newStudents, ...existingStudents];

            if (allStudentsToWrite.length === 0) {
                console.log("ImportFromExtension: No students to import");
                return;
            }

            console.log(`ImportFromExtension: Preparing to write ${allStudentsToWrite.length} students...`);
            const dataToWrite = [];
            const formulasToWrite = [];
            const cellsToColor = [];
            let gradebookLinksPreservedCount = 0;
            let assignedUsersPreservedCount = 0;

            allStudentsToWrite.forEach((incomingRow, index) => {
                const newRow = new Array(masterHeaders.length).fill("");
                const formulaRow = new Array(masterHeaders.length).fill(null);

                // Map incoming data to master list columns
                for (let incomingColIdx = 0; incomingColIdx < incomingRow.length; incomingColIdx++) {
                    const masterColIdx = colMapping[incomingColIdx];
                    if (masterColIdx !== -1) {
                        let cellValue = incomingRow[incomingColIdx] ?? "";

                        // Format student name to "Last, First"
                        if (masterColIdx === masterStudentNameCol) {
                            cellValue = formatToLastFirst(String(cellValue));
                        }

                        // Wrap Gradebook URLs in HYPERLINK formula
                        if (masterColIdx === masterGradebookCol && cellValue) {
                            const urlString = String(cellValue).trim();
                            // Check if it looks like a URL
                            if (urlString.startsWith('http://') || urlString.startsWith('https://')) {
                                formulaRow[masterColIdx] = `=HYPERLINK("${urlString}", "Grade Book")`;
                                cellValue = "Grade Book"; // Display value
                            }
                        }

                        newRow[masterColIdx] = cellValue;
                    }
                }

                // Handle missing assignments count based on gradebook link
                if (masterMissingAssignmentsCol !== -1) {
                    const incomingMissingCount = incomingMissingAssignmentsCol !== -1 ? incomingRow[incomingMissingAssignmentsCol] : null;
                    const hasGradebookLink = masterGradebookCol !== -1 && newRow[masterGradebookCol];

                    if (hasGradebookLink) {
                        // If gradebook link exists, default to 0 if no count provided
                        const missingCount = (incomingMissingCount !== null &&
                                            incomingMissingCount !== undefined &&
                                            incomingMissingCount !== '')
                                            ? incomingMissingCount
                                            : 0;
                        newRow[masterMissingAssignmentsCol] = missingCount;
                    } else {
                        // If no gradebook link, set to blank
                        newRow[masterMissingAssignmentsCol] = '';
                    }
                }

                // Preserve existing data for this student if they already exist
                const studentName = incomingRow[incomingStudentNameCol];
                const normalizedName = normalizeName(studentName);
                if (masterDataMap.has(normalizedName)) {
                    const existingData = masterDataMap.get(normalizedName);

                    // Preserve Gradebook links
                    if (existingData.gradebookFormula) {
                        if (masterGradebookCol !== -1 && !newRow[masterGradebookCol]) {
                            formulaRow[masterGradebookCol] = existingData.gradebookFormula;
                            const match = existingData.gradebookFormula.match(/, *"([^"]+)"\)/i);
                            newRow[masterGradebookCol] = match ? match[1] : "Gradebook";
                            gradebookLinksPreservedCount++;
                        }
                    }

                    // Preserve Assigned values
                    if (existingData.assigned) {
                        if (masterAssignedCol !== -1 && !newRow[masterAssignedCol]) {
                            newRow[masterAssignedCol] = existingData.assigned;
                            assignedUsersPreservedCount++;
                        }
                    }
                }

                // Preserve values from unmatched master columns
                if (nonBlankUnmatchedCols.length > 0) {
                    let preserveKey;
                    if (incomingIdCol !== -1) {
                        const idValue = incomingRow[incomingIdCol];
                        if (idValue != null && String(idValue).trim() !== '') {
                            preserveKey = String(idValue).trim();
                        }
                    }
                    if (!preserveKey) {
                        preserveKey = normalizedName;
                    }

                    if (preserveKey && preservationMap.has(preserveKey)) {
                        const preserved = preservationMap.get(preserveKey);
                        for (const colIdx of nonBlankUnmatchedCols) {
                            if (preserved.formulas[colIdx]) {
                                formulaRow[colIdx] = preserved.formulas[colIdx];
                                newRow[colIdx] = preserved.values[colIdx] || "";
                            } else if (preserved.values[colIdx] !== undefined) {
                                newRow[colIdx] = preserved.values[colIdx];
                            }
                        }
                    }
                }

                // Check for preserved colors based on Assigned column value
                if (masterAssignedCol !== -1) {
                    const assignedValue = newRow[masterAssignedCol];
                    if (assignedValue && valueToColorMap.has(assignedValue)) {
                        cellsToColor.push({
                            rowIndex: index + 1,
                            colIndex: masterAssignedCol,
                            color: valueToColorMap.get(assignedValue)
                        });
                    }
                }

                dataToWrite.push(newRow);
                formulasToWrite.push(formulaRow);
            });

            if (gradebookLinksPreservedCount > 0) {
                console.log(`ImportFromExtension: Preserved ${gradebookLinksPreservedCount} Gradebook links`);
            }
            if (assignedUsersPreservedCount > 0) {
                console.log(`ImportFromExtension: Preserved ${assignedUsersPreservedCount} Assigned users`);
            }

            // Convert Excel serial date numbers in Course Start/End columns before writing
            await applyCourseDateFormatting(context, sheet, masterHeaders, dataToWrite);

            // Write data and formulas in batches to avoid payload size limits
            const totalBatches = Math.ceil(dataToWrite.length / BATCH_SIZE);
            console.log(`ImportFromExtension: Writing ${dataToWrite.length} rows in ${totalBatches} batch(es) of up to ${BATCH_SIZE} rows...`);
            sendImportStatus('writing_data', `Writing ${dataToWrite.length} students...`, 0);

            for (let batchStart = 0; batchStart < dataToWrite.length; batchStart += BATCH_SIZE) {
                const batchEnd = Math.min(batchStart + BATCH_SIZE, dataToWrite.length);
                const batchSize = batchEnd - batchStart;

                const batchValues = dataToWrite.slice(batchStart, batchEnd);
                const batchFormulas = formulasToWrite.slice(batchStart, batchEnd);

                const batchRange = sheet.getRangeByIndexes(1 + batchStart, 0, batchSize, masterHeaders.length);
                batchRange.values = batchValues;
                batchRange.formulas = batchFormulas;

                await context.sync();

                const currentBatch = Math.floor(batchStart / BATCH_SIZE) + 1;
                console.log(`ImportFromExtension: Data write batch ${currentBatch}/${totalBatches} completed (rows ${batchStart + 1}-${batchEnd})`);
                sendImportStatus('writing_data', `Writing batch ${currentBatch} of ${totalBatches}...`, Math.round((currentBatch / totalBatches) * 100));
            }
            console.log("ImportFromExtension: All data writes completed");

            sendImportStatus('highlighting', 'Highlighting new students and preserved data...');

            // Highlight preserved columns in light gray (applied first so cell-level colors can override)
            if (nonBlankUnmatchedCols.length > 0) {
                for (const colIdx of nonBlankUnmatchedCols) {
                    const colRange = sheet.getRangeByIndexes(1, colIdx, dataToWrite.length, 1);
                    colRange.format.fill.color = "#EDEDED"; // Lighter Gray
                }
                console.log(`ImportFromExtension: Highlighted ${nonBlankUnmatchedCols.length} preserved column(s) in light gray: [${nonBlankUnmatchedCols.map(i => masterHeaders[i]).join(', ')}]`);
            }

            // Highlight new students based on the latest ExpStartDate
            const masterExpStartDateCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.expectedStartDate);
            if (masterExpStartDateCol !== -1) {
                // Find the latest ExpStartDate across all students
                let latestDate = null;
                for (let i = 0; i < dataToWrite.length; i++) {
                    const dateVal = parseDate(dataToWrite[i][masterExpStartDateCol]);
                    if (dateVal && (!latestDate || dateVal > latestDate)) {
                        latestDate = dateVal;
                    }
                }

                if (latestDate) {
                    const latestDateStr = latestDate.toDateString();
                    let highlightedCount = 0;
                    for (let i = 0; i < dataToWrite.length; i++) {
                        const dateVal = parseDate(dataToWrite[i][masterExpStartDateCol]);
                        if (dateVal && dateVal.toDateString() === latestDateStr) {
                            const rowRange = sheet.getRangeByIndexes(i + 1, 0, 1, masterHeaders.length);
                            rowRange.format.fill.color = "#ADD8E6"; // Light Blue
                            highlightedCount++;
                        }
                    }
                    console.log(`ImportFromExtension: Highlighted ${highlightedCount} students with latest ExpStartDate (${latestDateStr})`);
                }
            } else if (newStudents.length > 0) {
                // Fallback: highlight new students if ExpStartDate column doesn't exist
                console.log(`ImportFromExtension: Highlighting ${newStudents.length} new students...`);
                const highlightRange = sheet.getRangeByIndexes(1, 0, newStudents.length, masterHeaders.length);
                highlightRange.format.fill.color = "#ADD8E6"; // Light Blue
            }

            // Apply preserved cell colors last so they are not overwritten by column/row fills
            if (cellsToColor.length > 0) {
                const colorBatches = Math.ceil(cellsToColor.length / FORMAT_BATCH_SIZE);
                console.log(`ImportFromExtension: Applying ${cellsToColor.length} preserved cell colors in ${colorBatches} batch(es)...`);

                for (let i = 0; i < cellsToColor.length; i += FORMAT_BATCH_SIZE) {
                    const batchEnd = Math.min(i + FORMAT_BATCH_SIZE, cellsToColor.length);
                    for (let j = i; j < batchEnd; j++) {
                        const cell = cellsToColor[j];
                        sheet.getCell(cell.rowIndex, cell.colIndex).format.fill.color = cell.color;
                    }
                    await context.sync();
                }
            }

            // Apply conditional formatting to columns
            sendImportStatus('formatting', 'Applying conditional formatting...');
            await applyGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyLastCourseGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyMissingAssignmentsConditionalFormatting(context, sheet, masterHeaders);
            await applyHoldConditionalFormatting(context, sheet, masterHeaders);
            await applyAdSAPStatusConditionalFormatting(context, sheet, masterHeaders);
            await applyNextAssignmentDueFormatting(context, sheet, masterHeaders);
            await applyAttendanceConditionalFormatting(context, sheet, masterHeaders);
            await applyLetterGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyEnrollGpaConditionalFormatting(context, sheet, masterHeaders);

            // Update workbook settings with new columns for LDA column selector
            if (newColumns.length > 0) {
                try {
                    const docSettings = Office.context.document.settings.get('workbookSettings');
                    if (docSettings && Array.isArray(docSettings.columns)) {
                        const normalize = (s) => String(s || '').toLowerCase().replace(/\s+/g, '');
                        const existingSet = new Set(docSettings.columns.map(c => normalize(c.name)));
                        for (const colName of newColumns) {
                            if (!existingSet.has(normalize(colName))) {
                                docSettings.columns.push({ name: colName });
                            }
                        }
                        Office.context.document.settings.set('workbookSettings', docSettings);
                        Office.context.document.settings.saveAsync();
                        console.log('ImportFromExtension: Updated workbook settings with new columns:', newColumns);
                    }
                } catch (e) {
                    console.warn('ImportFromExtension: Failed to update workbook settings:', e);
                }
            }

            // Autofit columns
            sendImportStatus('autofitting', 'Auto-fitting columns...');
            console.log("ImportFromExtension: Autofitting columns...");
            sheet.getUsedRange().format.autofitColumns();

            await context.sync();
            console.log("ImportFromExtension: Master List import completed successfully");
        });

        // Check if the payload includes student objects with missing assignments
        if (payload.students && Array.isArray(payload.students)) {
            const studentsWithAssignments = payload.students.filter(
                student => student.missingAssignments && Array.isArray(student.missingAssignments) && student.missingAssignments.length > 0
            );

            if (studentsWithAssignments.length > 0) {
                console.log(`ImportFromExtension: Found ${studentsWithAssignments.length} students with missing assignments, importing...`);
                sendImportStatus('missing_assignments', `Importing missing assignments for ${studentsWithAssignments.length} students...`);
                await importMissingAssignments(studentsWithAssignments);
            }
        }

        sendImportStatus('complete', 'Master list imported successfully!');

    } catch (error) {
        console.error("ImportFromExtension: Error importing Master List:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("ImportFromExtension: Debug info:", JSON.stringify(error.debugInfo));
        }
        sendImportStatus('error', 'Import failed', 0, error.message || 'An unexpected error occurred');
    }
}

/**
 * Applies a 3-color scale conditional formatting to the grade column
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyGradeConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Grade column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const gradeColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);

        if (gradeColIdx === -1) {
            console.log("ImportFromExtension: Grade column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const gradeColumnRange = sheet.getRangeByIndexes(1, gradeColIdx, range.rowCount - 1, 1);

        // Determine if grades are 0-1 or 0-100 scale by checking the first few values
        let isPercentScale = false;
        for (let i = 1; i < Math.min(range.rowCount, 10); i++) {
            if (range.values[i] && typeof range.values[i][gradeColIdx] === 'number' && range.values[i][gradeColIdx] > 1) {
                isPercentScale = true;
                break;
            }
        }

        console.log(`ImportFromExtension: Detected grade scale: ${isPercentScale ? '0-100' : '0-1'}`);

        // Clear existing conditional formats on the column to avoid duplicates
        gradeColumnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Red (low) -> Yellow (mid) -> Green (high)
        const conditionalFormat = gradeColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" }, // Red
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: isPercentScale ? "70" : "0.7", color: "#FFEB84" }, // Yellow
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" } // Green
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Grade column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies a 3-color scale conditional formatting to the Last Course Grade column
 * (same Red-Yellow-Green scale as the Grade column)
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyLastCourseGradeConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Last Course Grade column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const colIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.lastCourseGrade);

        if (colIdx === -1) {
            console.log("ImportFromExtension: Last Course Grade column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const columnRange = sheet.getRangeByIndexes(1, colIdx, range.rowCount - 1, 1);

        // Determine if grades are 0-1 or 0-100 scale by checking the first few values
        let isPercentScale = false;
        for (let i = 1; i < Math.min(range.rowCount, 10); i++) {
            if (range.values[i] && typeof range.values[i][colIdx] === 'number' && range.values[i][colIdx] > 1) {
                isPercentScale = true;
                break;
            }
        }

        console.log(`ImportFromExtension: Detected Last Course Grade scale: ${isPercentScale ? '0-100' : '0-1'}`);

        // Clear existing conditional formats on the column to avoid duplicates
        columnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Red (low) -> Yellow (mid) -> Green (high)
        const conditionalFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: isPercentScale ? "70" : "0.7", color: "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Last Course Grade column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying Last Course Grade conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to the missing assignments column to highlight 0s in light green
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyMissingAssignmentsConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Missing Assignments column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const missingAssignmentsColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);

        if (missingAssignmentsColIdx === -1) {
            console.log("ImportFromExtension: Missing Assignments column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const missingAssignmentsColumnRange = sheet.getRangeByIndexes(1, missingAssignmentsColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        missingAssignmentsColumnRange.conditionalFormats.clearAll();

        // Apply conditional formatting: cells with value 0 get Green, Accent 6, Lighter 80% background
        const conditionalFormat = missingAssignmentsColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        conditionalFormat.cellValue.format.fill.color = "#E2EFDA"; // Green, Accent 6, Lighter 80%
        conditionalFormat.cellValue.rule = { formula1: "0", operator: "EqualTo" };

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Missing Assignments column (0s highlighted in Green, Accent 6, Lighter 80%)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying missing assignments conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to the Hold column to highlight "Yes" values in light red
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyHoldConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Hold column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const holdColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.hold);

        if (holdColIdx === -1) {
            console.log("ImportFromExtension: Hold column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const holdColumnRange = sheet.getRangeByIndexes(1, holdColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        holdColumnRange.conditionalFormats.clearAll();

        // Apply conditional formatting: cells with "Yes" get light red background
        const conditionalFormat = holdColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        conditionalFormat.textComparison.format.fill.color = "#FFB6C1"; // Light red (light pink)
        conditionalFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Yes" };

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Hold column ('Yes' highlighted in light red)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying hold conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to the AdSAPStatus column to highlight cells containing "Financial" in light red
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyAdSAPStatusConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to AdSAPStatus column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const adsapStatusColIdx = lowerCaseHeaders.indexOf('adsapstatus');

        if (adsapStatusColIdx === -1) {
            console.log("ImportFromExtension: AdSAPStatus column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const adsapStatusColumnRange = sheet.getRangeByIndexes(1, adsapStatusColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        adsapStatusColumnRange.conditionalFormats.clearAll();

        // Apply conditional formatting: cells containing "Financial" get light red background
        const conditionalFormat = adsapStatusColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        conditionalFormat.textComparison.format.fill.color = "#FFB6C1"; // Light red (light pink)
        conditionalFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Financial" };

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to AdSAPStatus column (cells containing 'Financial' highlighted in light red)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying AdSAPStatus conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies left text alignment to the Next Assignment Due column
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyNextAssignmentDueFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying left alignment to Next Assignment Due column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const nextAssignmentDueColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.nextAssignmentDue);

        if (nextAssignmentDueColIdx === -1) {
            console.log("ImportFromExtension: Next Assignment Due column not found, skipping formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        // Apply left alignment to the entire column (header + data)
        const columnRange = sheet.getRangeByIndexes(0, nextAssignmentDueColIdx, range.rowCount, 1);
        columnRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

        await context.sync();
        console.log("ImportFromExtension: Left alignment applied to Next Assignment Due column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying Next Assignment Due formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting and percentage number format to the Attendance % column.
 * 3-color scale: Red (lowest) -> Yellow (70%) -> Green (highest)
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyAttendanceConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Attendance % column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const attendanceColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.attendance);

        if (attendanceColIdx === -1) {
            console.log("ImportFromExtension: Attendance % column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const attendanceColumnRange = sheet.getRangeByIndexes(1, attendanceColIdx, range.rowCount - 1, 1);

        // Format as percentage so 0.38 displays as "38%"
        attendanceColumnRange.numberFormat = [["0%"]];

        // Clear existing conditional formats on the column to avoid duplicates
        attendanceColumnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Red (low) -> Yellow (70%) -> Green (high)
        const conditionalFormat = attendanceColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "0.7", color: "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting and percentage format applied to Attendance % column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying attendance conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to Letter Grade and Last Course Letter Grade columns.
 * Highlights cells beginning with "D" in light red and cells beginning with "F" in a darker red.
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyLetterGradeConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to letter grade columns...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const letterGradeColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.letterGrade);
        const lastCourseLetterGradeColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.lastCourseLetterGrade);

        const colIndices = [];
        if (letterGradeColIdx !== -1) colIndices.push({ idx: letterGradeColIdx, name: "Letter Grade" });
        if (lastCourseLetterGradeColIdx !== -1) colIndices.push({ idx: lastCourseLetterGradeColIdx, name: "Last Course Letter Grade" });

        if (colIndices.length === 0) {
            console.log("ImportFromExtension: No letter grade columns found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        for (const col of colIndices) {
            const columnRange = sheet.getRangeByIndexes(1, col.idx, range.rowCount - 1, 1);

            // Clear existing conditional formats on the column to avoid duplicates
            columnRange.conditionalFormats.clearAll();

            // Highlight cells beginning with "F" (darker red) - added first so it has lower priority
            const fFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
            fFormat.textComparison.format.fill.color = "#FF6B6B";
            fFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.beginsWith, text: "F" };

            // Highlight cells beginning with "D" (light red) - added second so it has higher priority
            const dFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
            dFormat.textComparison.format.fill.color = "#FFB6C1";
            dFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.beginsWith, text: "D" };

            console.log(`ImportFromExtension: Conditional formatting applied to ${col.name} column (D=light red, F=darker red)`);
        }

        await context.sync();
    } catch (error) {
        console.error("ImportFromExtension: Error applying letter grade conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies a 3-color scale conditional formatting to the Enroll GPA column.
 * Pink (0) -> Baby Blue (2) -> Light Green (4)
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
async function applyEnrollGpaConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Enroll GPA column...");

        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const enrollGpaColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.enrollGpa);

        if (enrollGpaColIdx === -1) {
            console.log("ImportFromExtension: Enroll GPA column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const columnRange = sheet.getRangeByIndexes(1, enrollGpaColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        columnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Pink (0) -> Blue (2) -> Green (4)
        const conditionalFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "0", color: "#FFC7CE" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "2", color: "#9BC2E6" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "4", color: "#C6EFCE" }
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Enroll GPA column (Pink #FFC7CE -> Blue #9BC2E6 -> Green #C6EFCE)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying Enroll GPA conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Converts Excel serial date numbers to M/DD/YYYY formatted strings in the
 * Course Start and Course End columns, then applies a date number format.
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 * @param {Array[]} dataToWrite The data rows (mutated in place)
 */
async function applyCourseDateFormatting(context, sheet, headers, dataToWrite) {
    try {
        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const courseStartColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.courseStart);
        const courseEndColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.courseEnd);

        const colIndices = [];
        if (courseStartColIdx !== -1) colIndices.push({ idx: courseStartColIdx, name: "Course Start" });
        if (courseEndColIdx !== -1) colIndices.push({ idx: courseEndColIdx, name: "Course End" });

        if (colIndices.length === 0) {
            return;
        }

        console.log(`ImportFromExtension: Formatting date columns: ${colIndices.map(c => c.name).join(', ')}`);

        // Convert Excel serial numbers to M/DD/YYYY strings in the data
        const excelSerialToDateString = (serial) => {
            if (typeof serial !== 'number' || serial < 1) return serial;
            // Excel serial: days since 1900-01-01 (with the 1900 leap year bug)
            const date = new Date((serial - 25569) * 86400 * 1000);
            if (isNaN(date.getTime())) return serial;
            const month = date.getUTCMonth() + 1;
            const day = date.getUTCDate();
            const year = date.getUTCFullYear();
            return `${month}/${String(day).padStart(2, '0')}/${year}`;
        };

        for (const col of colIndices) {
            let convertedCount = 0;
            for (let i = 0; i < dataToWrite.length; i++) {
                const val = dataToWrite[i][col.idx];
                if (typeof val === 'number' && val > 25569) {
                    dataToWrite[i][col.idx] = excelSerialToDateString(val);
                    convertedCount++;
                }
            }
            if (convertedCount > 0) {
                console.log(`ImportFromExtension: Converted ${convertedCount} serial dates in ${col.name} column`);
            }
        }
    } catch (error) {
        console.error("ImportFromExtension: Error formatting course date columns:", error);
    }
}

/**
 * Reads the Master List sheet and sends the data to the Chrome extension
 * @returns {Promise<object|null>} The Master List data as JSON, or null if sheet doesn't exist
 */
async function transferMasterList() {
    try {
        console.log("TransferMasterList: Starting to read Master List sheet...");

        const masterListData = await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Check if Master List sheet exists
            const masterListSheet = sheets.items.find(s => s.name === CONSTANTS.MASTER_LIST_SHEET);
            if (!masterListSheet) {
                console.warn("TransferMasterList: Master List sheet not found");
                return null;
            }

            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            // Load only dimensions first to avoid payload size limit
            usedRange.load("rowCount, columnCount");
            await context.sync();

            const mlTotalRows = usedRange.rowCount;
            const mlTotalCols = usedRange.columnCount;

            // Read values and formulas in batches to avoid response payload limits
            let mlValues = [];
            let mlFormulas = [];

            for (let startRow = 0; startRow < mlTotalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, mlTotalRows - startRow);
                const batchRange = sheet.getRangeByIndexes(startRow, 0, rowsToRead, mlTotalCols);
                batchRange.load("values, formulas");
                await context.sync();

                mlValues = mlValues.concat(batchRange.values);
                mlFormulas = mlFormulas.concat(batchRange.formulas);
            }

            // Parse headers - keep original case for headers array
            const rawHeaders = mlValues[0];
            const headers = rawHeaders.map(header => String(header || ''));
            const lowerCaseHeaders = headers.map(h => h.toLowerCase());

            // Find key column indices for reference
            const studentNameColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const gradeBookColIdx = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);

            // Create column mapping with actual header names
            const columnMapping = {};
            headers.forEach((header, idx) => {
                if (header) {
                    columnMapping[header] = idx;
                }
            });

            const students = [];
            const hyperlinkRegex = /=HYPERLINK\("([^"]+)"/i;

            // Process each row (skip header row)
            for (let i = 1; i < mlValues.length; i++) {
                const rowValues = mlValues[i];
                const rowFormulas = mlFormulas[i];

                // Only include row if it has a student name
                if (studentNameColIdx !== -1 && rowValues[studentNameColIdx]) {
                    const student = {};

                    // Add all columns dynamically
                    for (let colIdx = 0; colIdx < headers.length; colIdx++) {
                        const headerName = headers[colIdx];
                        if (!headerName) continue;

                        let value = rowValues[colIdx];

                        // Special handling for Gradebook column - extract URL from HYPERLINK
                        if (colIdx === gradeBookColIdx) {
                            const formula = rowFormulas[colIdx];
                            const match = String(formula).match(hyperlinkRegex);
                            if (match && match[1]) {
                                value = match[1]; // Extract URL from HYPERLINK formula
                            }
                        }

                        // Include value if it's not null/undefined/empty string
                        if (value !== null && value !== undefined && value !== "") {
                            student[headerName] = value;
                        }
                    }

                    students.push(student);
                }
            }

            console.log(`TransferMasterList: Sending ${headers.length} columns and ${students.length} students`);

            // Check if Missing Assignments sheet exists and parse it
            let missingAssignmentsData = null;
            try {
                const missingSheet = sheets.items.find(s => s.name === "Missing Assignments");
                if (missingSheet) {
                    console.log("TransferMasterList: Found Missing Assignments sheet, parsing data...");
                    const maSheet = context.workbook.worksheets.getItem("Missing Assignments");
                    const maRange = maSheet.getUsedRange();
                    // Load only dimensions first to avoid payload size limit
                    maRange.load("rowCount, columnCount");
                    await context.sync();

                    const maTotalRows = maRange.rowCount;
                    const maTotalCols = maRange.columnCount;

                    // Read values and formulas in batches
                    let maValues = [];
                    let maFormulasArr = [];

                    for (let startRow = 0; startRow < maTotalRows; startRow += BATCH_SIZE) {
                        const rowsToRead = Math.min(BATCH_SIZE, maTotalRows - startRow);
                        const batchRange = maSheet.getRangeByIndexes(startRow, 0, rowsToRead, maTotalCols);
                        batchRange.load("values, formulas");
                        await context.sync();

                        maValues = maValues.concat(batchRange.values);
                        maFormulasArr = maFormulasArr.concat(batchRange.formulas);
                    }

                    if (maValues.length > 0) {
                        // Parse headers from Missing Assignments sheet
                        const maHeaders = maValues[0].map(h => String(h || ''));
                        const maLowerHeaders = maHeaders.map(h => h.toLowerCase());

                        // Find the Gradebook column in Missing Assignments sheet
                        const maGradebookColIdx = findColumnIndex(maLowerHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);

                        console.log(`TransferMasterList: Missing Assignments headers: [${maHeaders.join(', ')}]`);
                        console.log(`TransferMasterList: Gradebook column index in Missing Assignments: ${maGradebookColIdx}`);

                        // Create a map of gradebook URL -> assignment data
                        const gradebookToAssignmentsMap = new Map();

                        for (let i = 1; i < maValues.length; i++) {
                            const rowValues = maValues[i];
                            const rowFormulas = maFormulasArr[i];

                            // Get the gradebook URL (extract from HYPERLINK if needed)
                            let gradebookUrl = null;
                            if (maGradebookColIdx !== -1) {
                                const formula = rowFormulas[maGradebookColIdx];
                                const match = String(formula).match(hyperlinkRegex);
                                if (match && match[1]) {
                                    gradebookUrl = match[1]; // Extract URL from HYPERLINK formula
                                } else {
                                    gradebookUrl = rowValues[maGradebookColIdx];
                                }
                            }

                            if (gradebookUrl) {
                                // Create an object with all columns from this row
                                // Use flattened structure that matches the import payload format
                                const assignmentData = {};
                                for (let colIdx = 0; colIdx < maHeaders.length; colIdx++) {
                                    const headerName = maHeaders[colIdx];
                                    if (!headerName) continue;

                                    let value = rowValues[colIdx];
                                    const formula = rowFormulas[colIdx];

                                    // Special handling for HYPERLINK formulas - extract URL and text separately
                                    const hyperlinkMatch = String(formula).match(/=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/i);
                                    if (hyperlinkMatch) {
                                        const url = hyperlinkMatch[1];
                                        const text = hyperlinkMatch[2];

                                        // Map to flattened properties based on column header
                                        if (headerName === "Assignment") {
                                            assignmentData.assignmentLink = url;
                                            assignmentData.assignmentTitle = text;
                                        } else if (headerName === "Submission") {
                                            assignmentData.submissionLink = url;
                                        } else if (headerName === "Grade Book") {
                                            // Grade Book is already handled separately, skip
                                        } else {
                                            // For any other hyperlink columns, store both url and text as separate properties
                                            assignmentData[`${headerName.toLowerCase().replace(/\s+/g, '')}Link`] = url;
                                            assignmentData[`${headerName.toLowerCase().replace(/\s+/g, '')}Text`] = text;
                                        }
                                    } else if (value !== null && value !== undefined && value !== "") {
                                        // Regular values - map to flattened property names
                                        if (headerName === "Due Date") {
                                            assignmentData.dueDate = value;
                                        } else if (headerName === "Score") {
                                            assignmentData.score = value;
                                        } else {
                                            // For any other columns, use the original header name
                                            assignmentData[headerName] = value;
                                        }
                                    }
                                }

                                // Store this assignment data keyed by gradebook URL
                                if (!gradebookToAssignmentsMap.has(gradebookUrl)) {
                                    gradebookToAssignmentsMap.set(gradebookUrl, []);
                                }
                                gradebookToAssignmentsMap.get(gradebookUrl).push(assignmentData);
                            }
                        }

                        console.log(`TransferMasterList: Created map with ${gradebookToAssignmentsMap.size} unique gradebook URLs from Missing Assignments`);

                        missingAssignmentsData = {
                            headers: maHeaders,
                            gradebookMap: gradebookToAssignmentsMap
                        };
                    }
                } else {
                    console.log("TransferMasterList: No Missing Assignments sheet found");
                }
            } catch (error) {
                console.error("TransferMasterList: Error reading Missing Assignments sheet:", error);
                // Continue without missing assignments data
            }

            // Add missing assignments to students if available
            if (missingAssignmentsData && missingAssignmentsData.gradebookMap.size > 0) {
                students.forEach(student => {
                    const studentGradebook = student[headers[gradeBookColIdx]]; // Get gradebook URL from student
                    if (studentGradebook && missingAssignmentsData.gradebookMap.has(studentGradebook)) {
                        student.missingAssignments = missingAssignmentsData.gradebookMap.get(studentGradebook);
                    }
                });
                console.log(`TransferMasterList: Added missing assignments data to students`);
            }

            return {
                sheetName: CONSTANTS.MASTER_LIST_SHEET,
                headers: headers,  // All column headers
                columnMapping: columnMapping,  // Map of header name -> column index
                students: students,  // Array of student objects with all columns
                totalStudents: students.length,
                timestamp: new Date().toISOString()
            };
        });

        if (masterListData) {
            console.log(`TransferMasterList: Successfully read ${masterListData.totalStudents} students from Master List with ${masterListData.headers.length} columns`);

            // Send the data to Chrome extension via postMessage
            chromeExtensionService.sendMessage({
                type: 'SRK_MASTER_LIST_DATA',
                data: masterListData
            });

            console.log("TransferMasterList: Data sent to Chrome extension");
        }

        return masterListData;

    } catch (error) {
        console.error("TransferMasterList: Error reading Master List:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info:", JSON.stringify(error.debugInfo));
        }
        return null;
    }
}

/**
 * Sets up listener to respond to ping checks from the Chrome extension
 * Call this during add-in initialization to enable connectivity checks
 */
function setupPingResponseListener() {
    window.addEventListener("message", (event) => {
        // Log all incoming messages for debugging
        if (event.data && event.data.type && event.data.type.startsWith('SRK_')) {
            console.log(`📨 Background Service: Received ${event.data.type} message:`, event.data);
        }

        if (event.data && event.data.type === "SRK_PING") {
            console.log('🏓 Background Service: Received SRK_PING, sending pong...');

            // Send pong response back to extension
            window.postMessage({
                type: "SRK_PONG",
                timestamp: new Date().toISOString(),
                source: "excel-addin-background"
            }, "*");

            console.log('✅ Background Service: SRK_PONG sent to Chrome extension');
        }

        if (event.data && event.data.type === "SRK_LINKS") {
            console.log('🔗 Background Service: Received SRK_LINKS message with links:', event.data.links);
            console.log('ℹ️ Background Service: SRK_LINKS is forwarded to Chrome extension (not handled here)');
        }
    });

    console.log('🔔 Background Service: Ping response listener set up');
}

/**
 * Creates a new sheet with the given name and headers
 * @param {string} sheetName - The name of the sheet to create
 * @param {string[]} headers - Array of column headers
 * @returns {Promise<void>}
 */
async function createSheetFromExtension(sheetName, headers) {
    try {
        console.log(`CreateSheet: Received request to create sheet "${sheetName}" with headers:`, headers);

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Check if sheet already exists
            let targetSheet = sheets.items.find(s => s.name === sheetName);

            if (!targetSheet) {
                console.log(`CreateSheet: Creating new sheet "${sheetName}"...`);
                targetSheet = sheets.add(sheetName);
                await context.sync();
            } else {
                console.log(`CreateSheet: Sheet "${sheetName}" already exists, getting reference...`);
                targetSheet = context.workbook.worksheets.getItem(sheetName);
            }

            // Write headers to the first row
            if (headers && Array.isArray(headers) && headers.length > 0) {
                const headerRange = targetSheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                await context.sync();
                console.log(`CreateSheet: Headers written to sheet "${sheetName}"`);
            }

            // Activate the new sheet
            targetSheet.activate();
            await context.sync();

            console.log(`CreateSheet: Sheet "${sheetName}" created and activated successfully`);
        });
    } catch (error) {
        console.error(`CreateSheet: Error creating sheet "${sheetName}":`, error);
        if (error instanceof OfficeExtension.Error) {
            console.error("CreateSheet: Debug info:", JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Retrieves all sheet names from the workbook and sends them to the Chrome extension
 * @returns {Promise<void>}
 */
async function sendSheetListToExtension() {
    try {
        console.log("SheetList: Retrieving all sheet names...");

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Extract sheet names
            const sheetNames = sheets.items.map(sheet => sheet.name);
            console.log(`SheetList: Found ${sheetNames.length} sheets:`, sheetNames);

            // Send response to Chrome extension
            chromeExtensionService.sendMessage({
                type: 'SRK_SHEET_LIST_RESPONSE',
                sheets: sheetNames,
                timestamp: new Date().toISOString()
            });

            console.log("SheetList: Sheet list sent to Chrome extension");
        });
    } catch (error) {
        console.error("SheetList: Error retrieving sheet names:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("SheetList: Debug info:", JSON.stringify(error.debugInfo));
        }

        // Send error response to extension
        chromeExtensionService.sendMessage({
            type: 'SRK_SHEET_LIST_RESPONSE',
            sheets: [],
            error: error.message,
            timestamp: new Date().toISOString()
        });
    }
}

/**
 * Sets up a listener on the "SRK_Commands" sheet to process highlight commands
 * published by Power Automate. When a command is written to the sheet via
 * co-authoring sync, the onChanged event fires and the add-in executes the
 * highlight locally in the user's session for instant visibility.
 *
 * If the SRK_Commands sheet doesn't exist yet, listens for new sheet creation
 * via the WorksheetCollection.onAdded event and attaches the listener when it appears.
 */
let commandQueueListenerAttached = false;

async function setupCommandQueueListener() {
    const COMMANDS_SHEET = "SRK_Commands";

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const cmdSheet = sheets.items.find(s => s.name === COMMANDS_SHEET);
            if (cmdSheet) {
                // Sheet exists — attach the onChanged listener
                await attachCommandSheetListener(context, COMMANDS_SHEET);
            } else {
                console.log("CommandQueue: SRK_Commands sheet not found. Watching for sheet creation...");
            }

            // Always listen for new sheets so we can attach when SRK_Commands is created
            if (!commandQueueListenerAttached) {
                sheets.onAdded.add(async (event) => {
                    try {
                        console.log(`CommandQueue: New sheet added: ${event.worksheetId}`);
                        // Check if the newly added sheet is our commands sheet
                        await Excel.run(async (ctx) => {
                            const newSheet = ctx.workbook.worksheets.getItem(event.worksheetId);
                            newSheet.load("name");
                            await ctx.sync();

                            if (newSheet.name === COMMANDS_SHEET) {
                                console.log("CommandQueue: SRK_Commands sheet detected! Attaching listener...");
                                await attachCommandSheetListener(ctx, COMMANDS_SHEET);
                            }
                        });
                    } catch (error) {
                        console.error("CommandQueue: Error handling new sheet event:", error);
                    }
                });
                commandQueueListenerAttached = true;
                await context.sync();
            }
        });
    } catch (error) {
        console.error("CommandQueue: Error setting up listener:", error);
    }
}

/**
 * Attaches the onChanged event handler to the SRK_Commands sheet.
 */
async function attachCommandSheetListener(context, sheetName) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);

    worksheet.onChanged.add(async (event) => {
        try {
            console.log(`CommandQueue: Change detected on ${sheetName} (source: ${event.source})`);
            await processCommandQueue();
        } catch (error) {
            console.error("CommandQueue: Error processing change event:", error);
        }
    });

    await context.sync();
    console.log("CommandQueue: Listening for highlight commands on SRK_Commands sheet");

    // Process any pending commands already in the sheet
    await processCommandQueue();
}

/**
 * Reads all pending commands from the SRK_Commands sheet, executes them,
 * and marks them as consumed by updating their status.
 */
async function processCommandQueue() {
    const COMMANDS_SHEET = "SRK_Commands";

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const cmdSheetExists = sheets.items.find(s => s.name === COMMANDS_SHEET);
            if (!cmdSheetExists) return;

            const sheet = context.workbook.worksheets.getItem(COMMANDS_SHEET);
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("values, rowCount, columnCount");
            await context.sync();

            if (usedRange.isNullObject || usedRange.rowCount <= 1) {
                return; // Only header row or empty
            }

            const values = usedRange.values;
            let processedCount = 0;

            for (let r = 1; r < values.length; r++) {
                const commandJson = String(values[r][0] || "").trim();
                const status = String(values[r][2] || "").trim();

                // Skip already processed or empty commands
                if (!commandJson || status === "done" || status === "error") {
                    continue;
                }

                try {
                    const command = JSON.parse(commandJson);

                    if (command.type === "SRK_HIGHLIGHT_STUDENT_ROW" && command.data) {
                        console.log(`CommandQueue: Executing highlight for student "${command.data.syStudentId}" on "${command.data.targetSheet}"`);

                        // Reuse the existing highlight handler from chromeExtensionService
                        // which runs in the user's local session for instant visibility
                        await chromeExtensionService.handleHighlightStudentRow(command.data);

                        // Mark as done
                        sheet.getCell(r, 2).values = [["done"]];
                        processedCount++;
                        console.log(`CommandQueue: Command row ${r + 1} executed successfully`);
                    } else {
                        console.warn(`CommandQueue: Unknown command type at row ${r + 1}:`, command.type);
                        sheet.getCell(r, 2).values = [["error"]];
                    }
                } catch (parseError) {
                    console.error(`CommandQueue: Failed to parse/execute command at row ${r + 1}:`, parseError);
                    sheet.getCell(r, 2).values = [["error"]];
                }
            }

            if (processedCount > 0) {
                await context.sync();
                console.log(`CommandQueue: Processed ${processedCount} command(s)`);
            }
        });
    } catch (error) {
        console.error("CommandQueue: Error processing queue:", error);
    }
}

/**
 * Polls the custom document property "SRK_Command" for highlight commands
 * written by Power Automate Office Scripts. Custom properties sync via
 * co-authoring much faster than cell values or formatting, enabling
 * near real-time highlights without a page refresh.
 *
 * Claim mechanism for multi-user:
 *   Each add-in instance has a unique session ID. When a pending command
 *   is detected, the instance writes its ID to "SRK_CommandClaim". After
 *   a brief delay, it re-reads the claim — if its ID is still there, it
 *   won the claim and executes. Otherwise, another instance handles it.
 *
 * Flow:
 *   1. Power Automate runs HighlightRows.ts which applies direct formatting
 *      AND writes the command to "SRK_Command" custom property
 *   2. Co-authoring syncs the property to all active sessions (~1-5 seconds)
 *   3. First add-in to claim the command executes the highlight locally
 *      (instant visibility), then clears the property
 */
const SRK_COMMAND_PROPERTY = "SRK_Command";
const SRK_CLAIM_PROPERTY = "SRK_CommandClaim";
const SRK_POLL_INTERVAL = 5000; // 5 seconds
const SRK_CLAIM_WAIT = 2000; // 2 seconds to wait before checking claim
const SRK_SESSION_ID = `session_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
let lastProcessedTimestamp = null;
let commandPollerActive = false;
let isProcessingCommand = false;

function startCommandPropertyPoller() {
    if (commandPollerActive) {
        console.log("CommandPropertyPoller: Already running, skipping duplicate start");
        return;
    }
    commandPollerActive = true;
    console.log(`CommandPropertyPoller: Session ID: ${SRK_SESSION_ID}`);
    console.log(`CommandPropertyPoller: Polling "${SRK_COMMAND_PROPERTY}" every ${SRK_POLL_INTERVAL / 1000}s`);

    setInterval(async () => {
        if (isProcessingCommand) return; // Don't overlap with an in-progress claim

        try {
            await Excel.run(async (context) => {
                const props = context.workbook.properties.custom;
                props.load("key, value");
                await context.sync();

                const srkProp = props.items.find(p => p.key === SRK_COMMAND_PROPERTY);
                if (!srkProp) return;

                const rawValue = srkProp.value;
                if (!rawValue || rawValue === "CLEAR") return;

                let command;
                try {
                    command = JSON.parse(rawValue);
                } catch (e) {
                    console.warn("CommandPropertyPoller: Invalid JSON, clearing:", rawValue);
                    context.workbook.properties.addCustomProperty(SRK_COMMAND_PROPERTY, "CLEAR");
                    await context.sync();
                    return;
                }

                // Skip if already processed this command
                if (command.timestamp && command.timestamp === lastProcessedTimestamp) {
                    return;
                }

                console.log(`CommandPropertyPoller: New command detected (timestamp: ${command.timestamp})`);

                // Try to claim the command
                isProcessingCommand = true;
                context.workbook.properties.addCustomProperty(SRK_CLAIM_PROPERTY, SRK_SESSION_ID);
                await context.sync();
            });

            if (!isProcessingCommand) return;

            // Wait briefly to let other sessions write their claim
            await new Promise(resolve => setTimeout(resolve, SRK_CLAIM_WAIT));

            // Re-read to see who won the claim
            await Excel.run(async (context) => {
                const props = context.workbook.properties.custom;
                props.load("key, value");
                await context.sync();

                const claimProp = props.items.find(p => p.key === SRK_CLAIM_PROPERTY);
                const claimValue = claimProp ? claimProp.value : null;

                // Re-read the command (it might have been cleared by the winner)
                const cmdProp = props.items.find(p => p.key === SRK_COMMAND_PROPERTY);
                const rawValue = cmdProp ? cmdProp.value : null;

                if (!rawValue || rawValue === "CLEAR") {
                    // Another session already processed and cleared it
                    console.log("CommandPropertyPoller: Command already handled by another session");
                    isProcessingCommand = false;
                    return;
                }

                let command;
                try {
                    command = JSON.parse(rawValue);
                } catch (e) {
                    isProcessingCommand = false;
                    return;
                }

                if (claimValue !== SRK_SESSION_ID) {
                    // Another session won the claim
                    console.log(`CommandPropertyPoller: Claim lost to ${claimValue}, skipping`);
                    lastProcessedTimestamp = command.timestamp;
                    isProcessingCommand = false;
                    return;
                }

                // We won the claim — execute the highlight
                console.log("CommandPropertyPoller: Claim won! Executing highlight...");

                if (command.type === "SRK_HIGHLIGHT_STUDENT_ROW" && command.data) {
                    try {
                        await chromeExtensionService.handleHighlightStudentRow(command.data);
                        console.log(`CommandPropertyPoller: Highlighted student "${command.data.syStudentId}" successfully`);
                    } catch (highlightError) {
                        console.error("CommandPropertyPoller: Highlight failed:", highlightError.message);
                    }
                } else {
                    console.warn("CommandPropertyPoller: Unknown command type:", command.type);
                }

                // Mark as processed and clear for next command
                lastProcessedTimestamp = command.timestamp;
                context.workbook.properties.addCustomProperty(SRK_COMMAND_PROPERTY, "CLEAR");
                context.workbook.properties.addCustomProperty(SRK_CLAIM_PROPERTY, "CLEAR");
                await context.sync();
                console.log("CommandPropertyPoller: Command completed, property cleared");
                isProcessingCommand = false;
            });
        } catch (error) {
            isProcessingCommand = false;
            if (error.code !== "GeneralException") {
                console.error("CommandPropertyPoller: Error:", error.message);
            }
        }
    }, SRK_POLL_INTERVAL);
}

// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("transferData", transferData);
Office.actions.associate("openAnalyticsPane", openAnalyticsPane);
Office.actions.associate("sendToCallQueue", createSendToCallQueue(chromeExtensionService));

// Register autoload event handler
Office.actions.associate("onDocumentOpen", onDocumentOpen);

// Initialize background services when Office is ready
Office.onReady(() => {
  console.log("Background script initialized - starting Chrome Extension Service");

  // Set up ping/pong listener for Chrome extension connectivity checks
  setupPingResponseListener();

  // Start extension detection
  // This will continuously ping to detect the Chrome extension
  chromeExtensionService.startPinging();

  // Start keep-alive heartbeat to prevent extension from going dormant
  chromeExtensionService.startKeepAlive();

  // Start listening for highlight commands from Power Automate
  setupCommandQueueListener();

  // Poll custom document property for highlight commands from Power Automate
  startCommandPropertyPoller();

  // Add a listener for extension events
  chromeExtensionService.addListener((event) => {
    console.log("Background: Chrome Extension event:", event);

    // Handle different event types
    if (event.type === "installed") {
      console.log("Background: Chrome extension detected and ready!");

      // When extension is detected, ask if it wants the Master List
      console.log("Background: Asking extension if it wants Master List data...");
      chromeExtensionService.sendMessage({
        type: 'SRK_REQUEST_MASTER_LIST',
        timestamp: new Date().toISOString()
      });
    } else if (event.type === "message") {
      console.log("Background: Received message from extension:", event.data);

      // Handle Master List request response
      if (event.data.type === 'SRK_MASTER_LIST_RESPONSE') {
        if (event.data.wantsData === true) {
          console.log("Background: Extension wants Master List data, sending...");
          transferMasterList();
        } else {
          console.log("Background: Extension declined Master List data");
        }
      }

      // Handle Master List import from extension
      if (event.data.type === 'SRK_IMPORT_MASTER_LIST') {
        console.log("Background: Received Master List import request from extension");
        importMasterListFromExtension(event.data.data);
      }

      // Handle sheet creation request from extension
      if (event.data.type === 'SRK_CREATE_SHEET') {
        console.log("Background: Received sheet creation request from extension");
        const { sheetName, headers } = event.data;
        if (sheetName && headers) {
          createSheetFromExtension(sheetName, headers);
        } else {
          console.error("Background: Invalid SRK_CREATE_SHEET request - missing sheetName or headers");
        }
      }

      // Handle sheet list request from extension
      if (event.data.type === 'SRK_REQUEST_SHEET_LIST') {
        console.log("Background: Received sheet list request from extension");
        sendSheetListToExtension();
      }
    }
  });

  console.log("Chrome Extension Service is now running in the background");
});
