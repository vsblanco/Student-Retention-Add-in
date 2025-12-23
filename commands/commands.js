/*
 * This file is the main entry point for all ribbon button commands.
 * It imports functions from other modules and associates them with Office actions.
 *
 * This file also serves as a background script that runs continuously,
 * even when taskpanes are closed, making it ideal for services that need
 * to always be listening (like the Chrome Extension Service).
 */
import { openImportDialog } from './import.js';
import { openCreateLdaDialog } from './lda.js';
import { toggleHighlight, transferData } from './actions.js';
import chromeExtensionService from '../react/src/services/chromeExtensionService.js';
import { CONSTANTS, findColumnIndex } from './utils.js';

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
            usedRange.load("values, formulas");
            await context.sync();

            // Parse headers
            const rawHeaders = usedRange.values[0];
            const headers = rawHeaders.map(header => String(header || '').toLowerCase());

            // Find column indices for all relevant columns
            const colIndices = {
                studentName: findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS),
                syStudentId: findColumnIndex(headers, CONSTANTS.STUDENT_ID_COLS),
                studentNumber: findColumnIndex(headers, CONSTANTS.STUDENT_NUMBER_COLS),
                gradeBook: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.gradeBook),
                daysOut: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.daysOut),
                lastLda: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.lastLda),
                grade: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.grade),
                primaryPhone: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.primaryPhone),
                otherPhone: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.otherPhone),
                personalEmail: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.PersonalEmail),
                studentEmail: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.StudentEmail),
                assigned: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.assigned),
                outreach: findColumnIndex(headers, CONSTANTS.OUTREACH_COLS)
            };

            const students = [];
            const hyperlinkRegex = /=HYPERLINK\("([^"]+)"/i;

            // Process each row (skip header row)
            for (let i = 1; i < usedRange.values.length; i++) {
                const rowValues = usedRange.values[i];
                const rowFormulas = usedRange.formulas[i];
                const student = {};
                let hasData = false;

                // Only include row if it has a student name
                if (colIndices.studentName !== -1 && rowValues[colIndices.studentName]) {
                    student.studentName = rowValues[colIndices.studentName];
                    hasData = true;

                    // Add all other fields if they exist
                    if (colIndices.syStudentId !== -1 && rowValues[colIndices.syStudentId]) {
                        student.syStudentId = rowValues[colIndices.syStudentId];
                    }
                    if (colIndices.studentNumber !== -1 && rowValues[colIndices.studentNumber]) {
                        student.studentNumber = rowValues[colIndices.studentNumber];
                    }
                    if (colIndices.gradeBook !== -1) {
                        const formula = rowFormulas[colIndices.gradeBook];
                        const match = String(formula).match(hyperlinkRegex);
                        if (match && match[1]) {
                            student.gradeBook = match[1]; // Extract URL from HYPERLINK formula
                        } else if (rowValues[colIndices.gradeBook]) {
                            student.gradeBook = rowValues[colIndices.gradeBook];
                        }
                    }
                    if (colIndices.daysOut !== -1 && rowValues[colIndices.daysOut] !== null && rowValues[colIndices.daysOut] !== "") {
                        student.daysOut = rowValues[colIndices.daysOut];
                    }
                    if (colIndices.lastLda !== -1 && rowValues[colIndices.lastLda]) {
                        student.lastLda = rowValues[colIndices.lastLda];
                    }
                    if (colIndices.grade !== -1 && rowValues[colIndices.grade] !== null && rowValues[colIndices.grade] !== "") {
                        student.grade = rowValues[colIndices.grade];
                    }
                    if (colIndices.primaryPhone !== -1 && rowValues[colIndices.primaryPhone]) {
                        student.primaryPhone = rowValues[colIndices.primaryPhone];
                    }
                    if (colIndices.otherPhone !== -1 && rowValues[colIndices.otherPhone]) {
                        student.otherPhone = rowValues[colIndices.otherPhone];
                    }
                    if (colIndices.personalEmail !== -1 && rowValues[colIndices.personalEmail]) {
                        student.personalEmail = rowValues[colIndices.personalEmail];
                    }
                    if (colIndices.studentEmail !== -1 && rowValues[colIndices.studentEmail]) {
                        student.studentEmail = rowValues[colIndices.studentEmail];
                    }
                    if (colIndices.assigned !== -1 && rowValues[colIndices.assigned]) {
                        student.assigned = rowValues[colIndices.assigned];
                    }
                    if (colIndices.outreach !== -1 && rowValues[colIndices.outreach]) {
                        student.outreach = rowValues[colIndices.outreach];
                    }
                }

                if (hasData) {
                    students.push(student);
                }
            }

            return {
                sheetName: CONSTANTS.MASTER_LIST_SHEET,
                columnMapping: colIndices,
                students: students,
                totalStudents: students.length,
                timestamp: new Date().toISOString()
            };
        });

        if (masterListData) {
            console.log(`TransferMasterList: Successfully read ${masterListData.totalStudents} students from Master List`);

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

// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("openImportDialog", openImportDialog);
Office.actions.associate("transferData", transferData);
Office.actions.associate("openCreateLdaDialog", openCreateLdaDialog);
Office.actions.associate("openAnalyticsPane", openAnalyticsPane);

// Register autoload event handler
Office.actions.associate("onDocumentOpen", onDocumentOpen);

// Initialize background services when Office is ready
Office.onReady(() => {
  console.log("Background script initialized - starting Chrome Extension Service");

  // Start extension detection
  // This will continuously ping to detect the Chrome extension
  chromeExtensionService.startPinging();

  // Start keep-alive heartbeat to prevent extension from going dormant
  chromeExtensionService.startKeepAlive();

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
    }
  });

  console.log("Chrome Extension Service is now running in the background");
});
