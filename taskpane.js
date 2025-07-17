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
 * Highlights the selected cell in yellow. If the cell is already yellow, it removes the highlight.
 * This function is called by a ribbon button.
 * @param {Office.AddinCommands.Event} event The event object passed from the ribbon button.
 */
async function toggleHighlight(event) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.load("color");
      await context.sync();

      if (range.format.fill.color === "#FFFF00") {
        range.format.fill.clear();
      } else {
        range.format.fill.color = "yellow";
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
  
  // Important: This tells the Office application that the function has completed its execution.
  if (event) {
    event.completed();
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
