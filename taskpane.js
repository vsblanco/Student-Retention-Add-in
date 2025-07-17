/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign click event handlers for the buttons in the taskpane.
    document.getElementById("insert-text-button").onclick = insertText;
    
    console.log("Add-in is ready to use in Excel!");
  }
});

/**
 * Finds the row of the current selection, and toggles a yellow highlight on the cells
 * in that row between the "StudentName" and "Outreach" columns.
 * This function assumes "StudentName" and "Outreach" are headers in the first row.
 * @param {Office.AddinCommands.Event} event The event object passed from the ribbon button.
 */
async function toggleHighlight(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selectedRange = context.workbook.getSelectedRange();
      
      // Load the row index of the user's selection to know which row to act on.
      selectedRange.load("rowIndex");
      
      // Get the first row of the sheet's used range to find the column headers.
      const headerRange = sheet.getUsedRange().getRow(0);
      headerRange.load("values");

      await context.sync();

      const headers = headerRange.values[0];
      const studentNameColIndex = headers.indexOf("StudentName");
      const outreachColIndex = headers.indexOf("Outreach");

      // Check if both header columns were found in the first row.
      if (studentNameColIndex === -1 || outreachColIndex === -1) {
        console.error("Error: Could not find 'StudentName' and/or 'Outreach' columns in the first row.");
        // In a real-world app, you might want to show a message to the user here.
        return; 
      }
      
      // Determine the start and end columns for the highlight range.
      const startCol = Math.min(studentNameColIndex, outreachColIndex);
      const endCol = Math.max(studentNameColIndex, outreachColIndex);
      const colCount = endCol - startCol + 1;

      // Get the specific range to highlight within the selected row.
      const targetRowIndex = selectedRange.rowIndex;
      const highlightRange = sheet.getRangeByIndexes(targetRowIndex, startCol, 1, colCount);
      
      // Load the fill color of the target range to check its current state.
      highlightRange.format.fill.load("color");
      await context.sync();

      // Toggle the highlight: if it's already yellow, clear it. Otherwise, make it yellow.
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
    // Important: This tells the Office application that the function has completed its execution.
    // It must be called regardless of whether the function succeeded or failed.
    if (event) {
      event.completed();
    }
  }
}

/**
 * Writes the string "Hello Excel!" into the currently selected cell in the worksheet.
 */
async function insertText() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.values = [["Hello Excel!"]];
      range.format.font.bold = true;
      await context.sync();
      console.log(`Successfully inserted text into cell: ${range.address}`);
    });
  } catch (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}

// This is a special function that Office calls to register ribbon button commands.
// It maps the function name from the manifest to the actual JavaScript function.
if (typeof Office.actions !== 'undefined') {
  Office.actions.associate("toggleHighlight", toggleHighlight);
}
