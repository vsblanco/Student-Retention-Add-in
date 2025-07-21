/**
 * This file contains the custom functions for the add-in.
 */

/**
 * Displays a given date, but shows today's date if the cell is green.
 * @customfunction LDA
 * @param {number} inputDate The date value from a cell.
 * @param {CustomFunctions.Invocation} invocation The invocation object from Excel.
 * @returns {Promise<number>} A promise that resolves to the date to be displayed.
 */
async function lda(inputDate, invocation) {
  // Get the address of the cell that invoked the function
  const address = invocation.address;
  let cellColor;

  try {
    await Excel.run(async (context) => {
      // Get the worksheet and the cell
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load("format/fill/color");
      await context.sync();
      cellColor = range.format.fill.color;
    });

    // Check if the color is a shade of green (you can adjust this logic)
    // This checks if the green component is the dominant one.
    if (cellColor && cellColor.startsWith("#")) {
      const r = parseInt(cellColor.substr(1, 2), 16);
      const g = parseInt(cellColor.substr(3, 2), 16);
      const b = parseInt(cellColor.substr(5, 2), 16);
      if (g > r && g > b) {
        // Return today's date as an Excel date serial number
        return excelDate(new Date());
      }
    }

    // If not green, or no color, return the original date
    return inputDate;

  } catch (error) {
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
