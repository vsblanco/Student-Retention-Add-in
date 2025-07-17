/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign a click event handler for the button.
    document.getElementById("insert-text-button").onclick = insertText;
    console.log("Add-in is ready to use in Excel!");
  }
});

/**
 * Writes the string "Hello Excel!" into the currently selected cell in the worksheet.
 */
async function insertText() {
  try {
    // Excel.run executes a batch of commands against the Excel API.
    // It handles creating the request context and other environment setup.
    await Excel.run(async (context) => {
      // Get the currently selected cell.
      const range = context.workbook.getSelectedRange();

      // Load the 'address' property of the range. 
      // This is not strictly necessary for writing, but it's good practice
      // to load properties you might need to read.
      range.load("address");

      // Update the range's value.
      range.values = [["Hello Excel!"]];

      // Change the font to bold.
      range.format.font.bold = true;

      // Synchronize the state between the script and the Office document.
      // This actually executes the queued commands (like setting the value).
      await context.sync();
      
      console.log(`Successfully inserted text into cell: ${range.address}`);
    });
  } catch (error) {
    // Always catch and log errors.
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}
