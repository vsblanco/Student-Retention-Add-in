/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let lastSelectedRow = -1; // Variable to track the last selected row index

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    setupTabs();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to register selection change handler: " + result.error.message);
      } else {
        console.log("Selection change handler registered successfully.");
      }
    });
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
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        panelDetails.classList.remove("hidden");
        panelHistory.classList.add("hidden");
    });

    tabHistory.addEventListener("click", () => {
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        panelHistory.classList.remove("hidden");
        panelDetails.classList.add("hidden");
    });
}

/**
 * Extracts initials from a name string. e.g., "Braddy, Diamond" -> "DB"
 * @param {string} name The student's name.
 * @returns {string} The initials.
 */
function getInitials(name) {
    if (!name || typeof name !== 'string') return '--';
    if (name.includes(',')) {
        const parts = name.split(',').map(part => part.trim());
        const lastName = parts[0];
        const firstName = parts[1];
        return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase();
    }
    const parts = name.split(' ').map(part => part.trim());
    if (parts.length > 1) {
        return `${parts[0].charAt(0)}${parts[parts.length - 1].charAt(0)}`.toUpperCase();
    }
    return name.substring(0, 2).toUpperCase();
}


/**
 * Handles the document selection change event.
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

            // Find all column indexes
            const colIdx = {
                name: lowerCaseHeaders.indexOf("studentname"),
                id: lowerCaseHeaders.indexOf("student id"), // Adjusted to common naming
                gender: lowerCaseHeaders.indexOf("gender"),
                daysOut: lowerCaseHeaders.indexOf("days out"),
                grade: lowerCaseHeaders.indexOf("grade"),
                status: lowerCaseHeaders.indexOf("status"),
                lastLda: lowerCaseHeaders.indexOf("last lda"),
                primaryPhone: lowerCaseHeaders.indexOf("primary phone"),
                otherPhone: lowerCaseHeaders.indexOf("other phone"),
                studentEmail: lowerCaseHeaders.indexOf("student email"),
                personalEmail: lowerCaseHeaders.indexOf("personal email"),
            };

            // Get all display elements
            const studentAvatar = document.getElementById("student-avatar");
            const studentNameDisplay = document.getElementById("student-name-display");
            const statusBadge = document.getElementById("status-badge");
            const studentIdDisplay = document.getElementById("student-id-display");
            const lastLdaDisplay = document.getElementById("last-lda-display");
            const daysOutDisplay = document.getElementById("days-out-display");
            const gradeDisplay = document.getElementById("grade-display");
            const gradeStatBlock = document.getElementById("grade-stat-block");
            const primaryPhoneDisplay = document.getElementById("primary-phone-display");
            const otherPhoneDisplay = document.getElementById("other-phone-display");
            const studentEmailDisplay = document.getElementById("student-email-display");
            const personalEmailDisplay = document.getElementById("personal-email-display");

            // Update UI
            const studentName = colIdx.name !== -1 ? rowData[colIdx.name] : "N/A";
            studentNameDisplay.textContent = studentName || "N/A";
            studentIdDisplay.textContent = (colIdx.id !== -1 ? rowData[colIdx.id] : "N/A") || "N/A";
            statusBadge.textContent = (colIdx.status !== -1 ? rowData[colIdx.status] : "N/A") || "N/A";
            lastLdaDisplay.textContent = (colIdx.lastLda !== -1 ? rowData[colIdx.lastLda] : "N/A") || "N/A";
            daysOutDisplay.textContent = (colIdx.daysOut !== -1 ? rowData[colIdx.daysOut] : "--") || "--";
            primaryPhoneDisplay.textContent = (colIdx.primaryPhone !== -1 ? rowData[colIdx.primaryPhone] : "N/A") || "N/A";
            otherPhoneDisplay.textContent = (colIdx.otherPhone !== -1 ? rowData[colIdx.otherPhone] : "N/A") || "N/A";
            studentEmailDisplay.textContent = (colIdx.studentEmail !== -1 ? rowData[colIdx.studentEmail] : "N/A") || "N/A";
            personalEmailDisplay.textContent = (colIdx.personalEmail !== -1 ? rowData[colIdx.personalEmail] : "N/A") || "N/A";

            // Avatar logic
            const gender = colIdx.gender !== -1 ? String(rowData[colIdx.gender]).toLowerCase() : "";
            studentAvatar.textContent = getInitials(studentName);
            if (gender === 'female') {
                studentAvatar.style.backgroundColor = '#ec4899'; // Pink
            } else if (gender === 'male') {
                studentAvatar.style.backgroundColor = '#3b82f6'; // Blue
            } else {
                studentAvatar.style.backgroundColor = '#6b7280'; // Gray
            }

            // Grade logic
            let grade = colIdx.grade !== -1 ? rowData[colIdx.grade] : null;
            gradeStatBlock.className = 'flex-1 p-3 text-center rounded-lg'; // Reset classes
            if (grade !== null && !isNaN(grade)) {
                const gradePercent = grade > 1 ? grade : grade * 100;
                gradeDisplay.textContent = `${Math.round(gradePercent)}%`;
                if (gradePercent >= 90) {
                    gradeStatBlock.classList.add('bg-green-200', 'text-green-800');
                } else if (gradePercent >= 70) {
                    gradeStatBlock.classList.add('bg-yellow-200', 'text-yellow-800');
                } else {
                    gradeStatBlock.classList.add('bg-red-200', 'text-red-800');
                }
            } else {
                gradeDisplay.textContent = 'N/A';
                gradeStatBlock.classList.add('bg-gray-200', 'text-gray-800');
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
