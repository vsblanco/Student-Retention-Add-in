/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let lastSelectedRow = -1; // Variable to track the last selected row index

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Register the event handler for selection changes.
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to register selection change handler: " + result.error.message);
      } else {
        console.log("Selection change handler registered successfully.");
      }
    });
    
    // Do an initial check on load.
    onSelectionChange();
  }
});

/**
 * Handles the document selection change event.
 * If the user selects a new row, it updates the task pane with the student's name from that row.
 */
async function onSelectionChange() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load("rowIndex");
            await context.sync();

            // Only proceed if the row has actually changed.
            if (selectedRange.rowIndex === lastSelectedRow) {
                return;
            }
            lastSelectedRow = selectedRange.rowIndex;

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const headerRange = sheet.getUsedRange().getRow(0);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0];
            const studentNameColIndex = headers.indexOf("StudentName");

            const studentNameDisplay = document.getElementById("student-name-display");

            // Check if the "StudentName" column exists.
            if (studentNameColIndex === -1) {
                studentNameDisplay.textContent = "N/A";
                console.error("Could not find 'StudentName' column in the first row.");
                return;
            }

            // Get the cell containing the student's name in the selected row.
            const nameCell = sheet.getRangeByIndexes(lastSelectedRow, studentNameColIndex, 1, 1);
            nameCell.load("values");
            await context.sync();

            const studentName = nameCell.values[0][0];
            studentNameDisplay.textContent = studentName || "Empty Cell"; // Display text if cell is empty
            studentNameDisplay.title = studentName || "Empty Cell";
        });
    } catch (error) {
        console.error("Error in onSelectionChange: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}


/**
 * Finds the row of the current selection, and toggles a yellow highlight on the cells
 * in that row between the "StudentName" and "Outreach" columns.
 * @param {Office.AddinCommands.Event} event The event object passed from the ribbon button.
 */
async function toggleHighlight(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("rowIndex");
      
      const headerRange = sheet.getUsedRange().getRow(0);
      headerRange.load("values");
      await context.sync();

      const headers = headerRange.values[0];
      const studentNameColIndex = headers.indexOf("StudentName");
      const outreachColIndex = headers.indexOf("Outreach");

      if (studentNameColIndex === -1 || outreachColIndex === -1) {
        console.error("Error: Could not find 'StudentName' and/or 'Outreach' columns in the first row.");
        return; 
      }
      
      const startCol = Math.min(studentNameColIndex, outreachColIndex);
      const endCol = Math.max(studentNameColIndex, outreachColIndex);
      const colCount = endCol - startCol + 1;

      const targetRowIndex = selectedRange.rowIndex;
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
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  } finally {
    if (event) {
      event.completed();
    }
  }
}

// This is a special function that Office calls to register ribbon button commands.
if (typeof Office.actions !== 'undefined') {
  Office.actions.associate("toggleHighlight", toggleHighlight);
}
