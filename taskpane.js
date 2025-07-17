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
 * If the user selects a new row, it updates the task pane with details from that row.
 * This version is optimized to use a single `context.sync()` call for better performance.
 */
async function onSelectionChange() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            // Load the row index and the entire used range of the sheet in one batch.
            selectedRange.load("rowIndex");
            const usedRange = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
            usedRange.load("values");

            // Execute the batch request.
            await context.sync();

            // Only proceed if the row has actually changed.
            if (selectedRange.rowIndex === lastSelectedRow) {
                return;
            }
            // If the selected row is the header row, do nothing.
            if (selectedRange.rowIndex === 0) {
                return;
            }
            lastSelectedRow = selectedRange.rowIndex;

            const sheetValues = usedRange.values;
            const headers = sheetValues[0];
            const rowData = sheetValues[lastSelectedRow];

            // Find the column indexes for the data we need.
            const studentNameColIndex = headers.indexOf("StudentName");
            const studentNumberColIndex = headers.indexOf("StudentNumber");
            const programVersionColIndex = headers.indexOf("ProgramVersion");
            const shiftColIndex = headers.indexOf("Shift");

            // Get the display elements from the DOM.
            const studentNameDisplay = document.getElementById("student-name-display");
            const studentNumberDisplay = document.getElementById("student-number-display");
            const programVersionDisplay = document.getElementById("program-version-display");
            const shiftDisplay = document.getElementById("shift-display");

            // Update Student Name
            const studentName = studentNameColIndex !== -1 ? rowData[studentNameColIndex] : "N/A";
            studentNameDisplay.textContent = studentName || "Empty Cell";
            studentNameDisplay.title = studentName || "Empty Cell";

            // Update Student Number
            const studentNumber = studentNumberColIndex !== -1 ? rowData[studentNumberColIndex] : "N/A";
            studentNumberDisplay.textContent = `Student #: ${studentNumber || 'N/A'}`;

            // Update Program Version with advanced parsing
            let programVersion = programVersionColIndex !== -1 ? rowData[programVersionColIndex] : "N/A";
            if (typeof programVersion === 'string' && programVersion !== 'N/A') {
                const match = programVersion.match(/\d{4}/); // Find the first 4-digit number (e.g., a year)
                if (match) {
                    programVersion = programVersion.substring(match.index + 4).trim();
                }
            }
            programVersionDisplay.textContent = programVersion || "N/A";
            
            // Update Shift
            const shift = shiftColIndex !== -1 ? rowData[shiftColIndex] : "N/A";
            shiftDisplay.textContent = shift || "N/A";
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
      const headerRange = sheet.getUsedRange().getRow(0);
      
      // Load properties in a single batch.
      selectedRange.load("rowIndex");
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
