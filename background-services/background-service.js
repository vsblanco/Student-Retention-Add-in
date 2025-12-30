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
                const grade = student.Grade || "";
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
                        row[5] = typeof assignment.Score === 'object' ? assignment.Score.text : assignment.score || assignment.Score || "";

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
                console.log(`ImportMissingAssignments: Writing ${dataToWrite.length} assignment rows...`);

                // Write data and formulas
                const dataRange = maSheet.getRangeByIndexes(1, 0, dataToWrite.length, headers.length);
                dataRange.values = dataToWrite;
                dataRange.formulas = formulasToWrite;

                // Autofit columns
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
            const masterMissingAssignmentsCol = findColumnIndex(lowerCaseMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const incomingMissingAssignmentsCol = findColumnIndex(lowerCaseIncomingHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);

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

            // Apply conditional formatting to columns
            await applyGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyMissingAssignmentsConditionalFormatting(context, sheet, masterHeaders);
            await applyHoldConditionalFormatting(context, sheet, masterHeaders);
            await applyAdSAPStatusConditionalFormatting(context, sheet, masterHeaders);

            // Autofit columns
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
                await importMissingAssignments(studentsWithAssignments);
            }
        }

    } catch (error) {
        console.error("ImportFromExtension: Error importing Master List:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("ImportFromExtension: Debug info:", JSON.stringify(error.debugInfo));
        }
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

            // Parse headers - keep original case for headers array
            const rawHeaders = usedRange.values[0];
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
            for (let i = 1; i < usedRange.values.length; i++) {
                const rowValues = usedRange.values[i];
                const rowFormulas = usedRange.formulas[i];

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
                    maRange.load("values, formulas");
                    await context.sync();

                    if (maRange.values.length > 0) {
                        // Parse headers from Missing Assignments sheet
                        const maHeaders = maRange.values[0].map(h => String(h || ''));
                        const maLowerHeaders = maHeaders.map(h => h.toLowerCase());

                        // Find the Gradebook column in Missing Assignments sheet
                        const maGradebookColIdx = findColumnIndex(maLowerHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);

                        console.log(`TransferMasterList: Missing Assignments headers: [${maHeaders.join(', ')}]`);
                        console.log(`TransferMasterList: Gradebook column index in Missing Assignments: ${maGradebookColIdx}`);

                        // Create a map of gradebook URL -> assignment data
                        const gradebookToAssignmentsMap = new Map();

                        for (let i = 1; i < maRange.values.length; i++) {
                            const rowValues = maRange.values[i];
                            const rowFormulas = maRange.formulas[i];

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
        if (event.data && event.data.type === "SRK_TASKPANE_PING") {
            console.log('🏓 Received ping from Chrome extension taskpane, sending pong...');

            // Send pong response back to extension
            window.postMessage({
                type: "SRK_TASKPANE_PONG",
                timestamp: new Date().toISOString()
            }, "*");

            console.log('✅ Pong sent to Chrome extension');
        }
    });

    console.log('🔔 Ping response listener set up');
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

  // Set up ping/pong listener for Chrome extension connectivity checks
  setupPingResponseListener();

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
