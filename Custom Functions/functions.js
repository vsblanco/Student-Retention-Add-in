/**
 * This file contains the custom functions for the add-in.
 * It includes logging to help with debugging.
 */

console.log("Custom functions script loaded successfully.");

/**
 * Displays a given date, but shows today's date if the cell is green.
 * @customfunction LDA
 * @param {number} inputDate The date value from a cell.
 * @param {CustomFunctions.Invocation} invocation The invocation object from Excel.
 * @returns {Promise<number>} A promise that resolves to the date to be displayed.
 */
async function lda(inputDate, invocation) {
  console.log("LDA function started.");
  // Get the address of the cell that invoked the function
  const address = invocation.address;
  console.log("Function invoked for cell: " + address);
  let cellColor;

  try {
    await Excel.run(async (context) => {
      // Get the worksheet and the cell
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load("format/fill/color");
      await context.sync();
      cellColor = range.format.fill.color;
      console.log("Cell color found: " + cellColor);
    });

    // Check if the color is a shade of green (you can adjust this logic)
    // This checks if the green component is the dominant one.
    if (cellColor && cellColor.startsWith("#")) {
      const r = parseInt(cellColor.substr(1, 2), 16);
      const g = parseInt(cellColor.substr(3, 2), 16);
      const b = parseInt(cellColor.substr(5, 2), 16);
      if (g > r && g > b) {
        const today = excelDate(new Date());
        console.log("Cell is green. Returning today's date: " + today);
        // Return today's date as an Excel date serial number
        return today;
      }
    }

    // If not green, or no color, return the original date
    console.log("Cell is not green. Returning original date: " + inputDate);
    return inputDate;

  } catch (error) {
    console.error("Error in LDA function: " + error.message);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
    return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Could not check cell color.");
  }
}

/**
 * Helper function to convert a JavaScript Date object to an Excel serial number.
 * @param {Date} date The date to convert.
 * @returns {number} The Excel date serial number.
 */
function excelDate(date) {
    return (date.getTime() / 86400000) + 25569;
}

// Register the custom function with Excel
CustomFunctions.associate("LDA", lda);
console.log("Custom function LDA associated.");
