/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let lastSelectedRow = -1; // Variable to track the last selected row index

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Setup Tab functionality
    setupTabs();
    
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
 * Sets up the event listeners for the tabbed interface.
 */
function setupTabs() {
    const tabDetails = document.getElementById("tab-details");
    const tabHistory = document.getElementById("tab-history");
    const panelDetails = document.getElementById("panel-details");
    const panelHistory = document.getElementById("panel-history");

    tabDetails.addEventListener("click", () => {
        // Style tabs
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        // Show/hide panels
        panelDetails.classList.remove("hidden");
        panelHistory.classList.add("hidden");
    });

    tabHistory.addEventListener("click", () => {
        // Style tabs
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        // Show/hide panels
        panelHistory.classList.remove("hidden");
        panelDetails.classList.add("hidden");
    });
}

/**
 * Handles the document selection change event.
 * If the user selects a new row, it updates the task pane with details from that row.
 */
async function onSelectionChange() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load("rowIndex");
            const usedRange = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
            usedRange.load("values");

            await context.sync();

            if (selectedRange.rowIndex === lastSelectedRow || selectedRange.rowIndex === 0) {
                return;
            }
            lastSelectedRow = selectedRange.rowIndex;

            const sheetValues = usedRange.values;
            const headers = sheetValues[0];
            const rowData = sheetValues[lastSelectedRow];

            const lowerCaseHeaders = headers.map(header => header.toLowerCase());

            const studentNameColIndex = lowerCaseHeaders.indexOf("studentname");
            const studentNumberColIndex = lowerCaseHeaders.indexOf("studentnumber");
            const programVersionColIndex = lowerCaseHeaders.indexOf("programversion");
            const shiftColIndex = lowerCaseHeaders.indexOf("shift");
            const gradeColIndex = lowerCaseHeaders.indexOf("grade");

            const studentNameDisplay = document.getElementById("student-name-display");
            const studentNumberDisplay = document.getElementById("student-number-display");
            const programVersionDisplay = document.getElementById("program-version-display");
            const shiftDisplay = document.getElementById("shift-display");
            const gradeDisplayBadge = document.getElementById("grade-display-badge");

            studentNameDisplay.textContent = (studentNameColIndex !== -1 ? rowData[studentNameColIndex] : "N/A") || "Empty Cell";
            studentNumberDisplay.textContent = `Student #: ${(studentNumberColIndex !== -1 ? rowData[studentNumberColIndex] : "N/A") || 'N/A'}`;
            shiftDisplay.textContent = (shiftColIndex !== -1 ? rowData[shiftColIndex] : "N/A") || "N/A";

            let programVersion = programVersionColIndex !== -1 ? rowData[programVersionColIndex] : "N/A";
            if (typeof programVersion === 'string' && programVersion !== 'N/A') {
                const match = programVersion.match(/\d{4}/);
                if (match) {
                    programVersion = programVersion.substring(match.index + 4).trim();
                }
            }
            programVersionDisplay.textContent = programVersion || "N/A";

            let grade = gradeColIndex !== -1 ? rowData[gradeColIndex] : null;
            gradeDisplayBadge.className = 'w-16 h-16 rounded-full flex items-center justify-center text-white font-bold text-xl';
            
            if (grade !== null && !isNaN(grade)) {
                const gradePercent = grade > 1 ? grade : grade * 100;
                gradeDisplayBadge.textContent = `${Math.round(gradePercent)}%`;
                if (gradePercent >= 90) {
                    gradeDisplayBadge.classList.add('bg-green-600');
                } else if (gradePercent >= 70) {
                    gradeDisplayBadge.classList.add('bg-yellow-500');
                } else {
                    gradeDisplayBadge.classList.add('bg-red-600');
                }
            } else {
                gradeDisplayBadge.textContent = 'N/A';
                gradeDisplayBadge.classList.add('bg-gray-400');
            }
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
      const lowerCaseHeaders = headers.map(header => header.toLowerCase());

      const studentNameColIndex = lowerCaseHeaders.indexOf("studentname");
      const outreachColIndex = lowerCaseHeaders.indexOf("outreach");

      if (studentNameColIndex === -1 || outreachColIndex === -1) {
        console.error("Could not find 'StudentName' and/or 'Outreach' columns.");
        return; 
      }
      
      const startCol = Math.min(studentNameColIndex, outreachColIndex);
      const endCol = Math.max(studentNameColIndex, outreachColIndex);
      const colCount = endCol - startCol + 1;
      const targetRowIndex = selectedRange.rowIndex;

      if (targetRowIndex === 0) return;

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

// Register ribbon button commands
if (typeof Office.actions !== 'undefined') {
  Office.actions.associate("toggleHighlight", toggleHighlight);
}
