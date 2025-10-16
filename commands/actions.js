/*
 * This file contains the logic for the "Contacted" and "Transfer Data" buttons.
 */
import { CONSTANTS, findColumnIndex } from './utils.js';

let transferDialog = null;

/**
 * Finds the row of the current selection, and toggles a yellow highlight on the cells
 * in that row between the "StudentName" and "Outreach" columns.
 */
export async function toggleHighlight(event) {
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
export async function transferData(event) {
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
                grade: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.grade),
                primaryPhone: findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.primaryPhone) // Added
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
                if (colIndices.primaryPhone !== -1 && rowValues[colIndices.primaryPhone]) {
                    rowData.PrimaryPhone = rowValues[colIndices.primaryPhone];
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
