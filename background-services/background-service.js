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
import { openImportDialog } from './data-import-handler.js';
import { toggleHighlight, transferData } from './ribbon-actions.js';
import chromeExtensionService from '../react/src/services/chromeExtensionService.js';
import { CONSTANTS, findColumnIndex, normalizeName, formatToLastFirst } from './shared-utilities.js';

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

        await Excel.run(async (context) => {
            // Check if Master List sheet exists
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const masterListSheet = sheets.items.find(s => s.name === CONSTANTS.MASTER_LIST_SHEET);
            if (!masterListSheet) {
                console.error("ImportFromExtension: Master List sheet not found. Import aborted.");
                return;
            }

            console.log("ImportFromExtension: Master List sheet found, proceeding with import...");

            // Get the Master List sheet
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, formulas, rowCount, rowIndex, columnIndex");
            await context.sync();

            // Get Master List headers
            const masterHeaders = usedRange.values[0].map(h => String(h || ''));
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

            console.log(`ImportFromExtension: Column mapping: [${colMapping.join(', ')}]`);

            // Find the student name column in both incoming data and master list
            const incomingStudentNameCol = findColumnIndex(lowerCaseIncomingHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterStudentNameCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterGradebookCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            const masterAssignedCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.assigned);

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
            for (let i = 1; i < usedRange.values.length; i++) {
                const name = usedRange.values[i][masterStudentNameCol];
                if (name) {
                    const normalizedName = normalizeName(name);
                    const gradebookFormula = (masterGradebookCol !== -1 && usedRange.formulas[i][masterGradebookCol]) ? usedRange.formulas[i][masterGradebookCol] : null;
                    const assignedValue = (masterAssignedCol !== -1) ? usedRange.values[i][masterAssignedCol] : null;

                    masterDataMap.set(normalizedName, {
                        gradebookFormula: gradebookFormula,
                        assigned: assignedValue
                    });
                }
            }

            console.log(`ImportFromExtension: Created map of ${masterDataMap.size} existing students`);

            // Get colors for assigned column values
            if (masterAssignedCol !== -1) {
                const allAssignedValues = usedRange.values.map(row => row[masterAssignedCol]);
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
            if (usedRange.rowCount > 1) {
                const rangeToClear = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, masterHeaders.length);
                rangeToClear.clear(Excel.ClearApplyTo.all);
                rangeToClear.getEntireRow().delete(Excel.DeleteShiftDirection.up);
                await context.sync();
                console.log("ImportFromExtension: Sheet cleared");
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
                        let cellValue = incomingRow[incomingColIdx] || "";
                        if (masterColIdx === masterStudentNameCol) {
                            cellValue = formatToLastFirst(String(cellValue));
                        }
                        newRow[masterColIdx] = cellValue;
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

            // Write all data and formulas
            console.log("ImportFromExtension: Writing data to sheet...");
            const writeRange = sheet.getRangeByIndexes(1, 0, dataToWrite.length, masterHeaders.length);
            writeRange.values = dataToWrite;
            writeRange.formulas = formulasToWrite;
            await context.sync();
            console.log("ImportFromExtension: Data write completed");

            // Apply preserved colors
            if (cellsToColor.length > 0) {
                console.log(`ImportFromExtension: Applying ${cellsToColor.length} preserved cell colors...`);
                for (const cell of cellsToColor) {
                    sheet.getCell(cell.rowIndex, cell.colIndex).format.fill.color = cell.color;
                }
            }

            // Highlight new students
            if (newStudents.length > 0) {
                console.log(`ImportFromExtension: Highlighting ${newStudents.length} new students...`);
                const highlightRange = sheet.getRangeByIndexes(1, 0, newStudents.length, masterHeaders.length);
                highlightRange.format.fill.color = "#ADD8E6"; // Light Blue
            }

            // Autofit columns
            console.log("ImportFromExtension: Autofitting columns...");
            sheet.getUsedRange().format.autofitColumns();

            await context.sync();
            console.log("ImportFromExtension: Master List import completed successfully");
        });

    } catch (error) {
        console.error("ImportFromExtension: Error importing Master List:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("ImportFromExtension: Debug info:", JSON.stringify(error.debugInfo));
        }
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

      // Handle Master List import from extension
      if (event.data.type === 'SRK_IMPORT_MASTER_LIST') {
        console.log("Background: Received Master List import request from extension");
        importMasterListFromExtension(event.data.data);
      }
    }
  });

  console.log("Chrome Extension Service is now running in the background");
});
