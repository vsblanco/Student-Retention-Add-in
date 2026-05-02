/*
 * master-list-transfer.js
 *
 * Reads the Master List sheet (and any companion Missing Assignments sheet),
 * normalizes it into a JSON payload, and sends it to the Chrome extension
 * via the chromeExtensionService bridge. Triggered when the extension
 * announces it wants the data.
 */
import { CONSTANTS, findColumnIndex } from '../shared-utilities.js';
import chromeExtensionService from '../../shared/chromeExtensionService.js';
import { BATCH_SIZE } from '../../shared/constants.js';
import { parseHyperlinkFormula } from '../../shared/excel-helpers.js';

/**
 * Reads the Master List sheet and sends the data to the Chrome extension
 * @returns {Promise<object|null>} The Master List data as JSON, or null if sheet doesn't exist
 */
export async function transferMasterList() {
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
                            const parsed = parseHyperlinkFormula(rowFormulas[colIdx]);
                            if (parsed) {
                                value = parsed.url;
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
                                const parsed = parseHyperlinkFormula(rowFormulas[maGradebookColIdx]);
                                gradebookUrl = parsed ? parsed.url : rowValues[maGradebookColIdx];
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
                                    const hyperlinkMatch = parseHyperlinkFormula(formula);
                                    if (hyperlinkMatch) {
                                        const { url, text } = hyperlinkMatch;

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
