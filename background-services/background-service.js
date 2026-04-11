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
