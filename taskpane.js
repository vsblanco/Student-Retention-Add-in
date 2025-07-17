/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign click event handlers for the buttons.
    document.getElementById("insert-text-button").onclick = insertText;
    document.getElementById("highlight-button").onclick = toggleHighlight;
    console.log("Add-in is ready to use in Excel!");
  }
});

/**
 * Highlights the selected cell in yellow. If the cell is already yellow, it removes the highlight.
 */
async function toggleHighlight() {
  try {
    await Excel.run(async (context) => {
      // Get the currently selected range.
      const range = context.workbook.getSelectedRange();
      
      // Load the 'color' property from the range's format/fill.
      range.format.fill.load("color");

      await context.sync();

      // Check the current color of the cell.
      // The color is returned in the format "#RRGGBB".
      if (range.format.fill.color === "#FFFF00") {
        // If the color is yellow, clear the fill to remove the highlight.
        range.format.fill.clear();
      } else {
        // If the color is not yellow, set it to yellow.
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
