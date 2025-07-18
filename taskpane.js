/*
 * This file contains the logic for the Task Pane UI.
 * It is loaded by taskpane.html.
 */

let lastSelectedRow = -1; // Variable to track the last selected row index

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // The DOM is not guaranteed to be ready when Office.onReady fires.
    // We will wait for the DOMContentLoaded event before running our setup code.
    document.addEventListener("DOMContentLoaded", function() {
        setupTabs();
        
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to register selection change handler: " + result.error.message);
          } else {
            console.log("Selection change handler registered successfully.");
          }
        });
        
        onSelectionChange();
    });
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
 * Finds the index of a column by checking against a list of possible names.
 */
function findColumnIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

/**
 * Handles the document selection change event.
 */
async function onSelectionChange() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load("rowIndex");
            await context.sync();

            if (selectedRange.rowIndex === lastSelectedRow) {
                return; // No change, exit early.
            }
            lastSelectedRow = selectedRange.rowIndex;

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Get the header row, assuming it's the first row of the used range.
            const headerRange = sheet.getUsedRange().getRow(0);
            headerRange.load(["values", "rowIndex", "columnCount"]);
            await context.sync();

            // Exit if the selection is the header row itself, or any row above it.
            if (selectedRange.rowIndex < headerRange.rowIndex + 1) {
                return;
            }

            // Get the specific row that was selected.
            const selectedRow = sheet.getRangeByIndexes(lastSelectedRow, 0, 1, headerRange.columnCount);
            selectedRow.load("values");
            await context.sync();

            const headers = headerRange.values[0];
            const rowData = selectedRow.values[0];
            
            const lowerCaseHeaders = headers.map(header => String(header || '').toLowerCase());

            const columnMappings = {
                name: ["studentname", "student name"],
                id: ["student id", "id"],
                gender: ["gender"],
                daysOut: ["days out", "daysout"],
                grade: ["grade", "course grade"],
                status: ["status"],
                lastLda: ["last lda", "lda"],
                primaryPhone: ["primary phone", "phone"],
                otherPhone: ["other phone", "cell phone", "cell","OtherPhone"],
                studentEmail: ["student email", "school email","email"],
                personalEmail: ["personal email", "otheremail"],
            };

            const colIdx = {
                name: findColumnIndex(lowerCaseHeaders, columnMappings.name),
                id: findColumnIndex(lowerCaseHeaders, columnMappings.id),
                gender: findColumnIndex(lowerCaseHeaders, columnMappings.gender),
                daysOut: findColumnIndex(lowerCaseHeaders, columnMappings.daysOut),
                grade: findColumnIndex(lowerCaseHeaders, columnMappings.grade),
                status: findColumnIndex(lowerCaseHeaders, columnMappings.status),
                lastLda: findColumnIndex(lowerCaseHeaders, columnMappings.lastLda),
                primaryPhone: findColumnIndex(lowerCaseHeaders, columnMappings.primaryPhone),
                otherPhone: findColumnIndex(lowerCaseHeaders, columnMappings.otherPhone),
                studentEmail: findColumnIndex(lowerCaseHeaders, columnMappings.studentEmail),
                personalEmail: findColumnIndex(lowerCaseHeaders, columnMappings.personalEmail),
            };

            const studentAvatar = document.getElementById("student-avatar");
            const studentNameDisplay = document.getElementById("student-name-display");
            const statusBadge = document.getElementById("status-badge");
            const studentIdDisplay = document.getElementById("student-id-display");
            const lastLdaDisplay = document.getElementById("last-lda-display");
            const daysOutDisplay = document.getElementById("days-out-display");
            const daysOutStatBlock = document.getElementById("days-out-stat-block");
            const gradeDisplay = document.getElementById("grade-display");
            const gradeStatBlock = document.getElementById("grade-stat-block");
            const primaryPhoneDisplay = document.getElementById("primary-phone-display");
            const otherPhoneDisplay = document.getElementById("other-phone-display");
            const studentEmailDisplay = document.getElementById("student-email-display");
            const personalEmailDisplay = document.getElementById("personal-email-display");

            const studentName = colIdx.name !== -1 ? rowData[colIdx.name] : "N/A";
            studentNameDisplay.textContent = studentName || "N/A";
            studentIdDisplay.textContent = (colIdx.id !== -1 ? rowData[colIdx.id] : "N/A") || "N/A";
            statusBadge.textContent = (colIdx.status !== -1 ? rowData[colIdx.status] : "N/A") || "N/A";
            lastLdaDisplay.textContent = (colIdx.lastLda !== -1 ? rowData[colIdx.lastLda] : "N/A") || "N/A";
            primaryPhoneDisplay.textContent = (colIdx.primaryPhone !== -1 ? rowData[colIdx.primaryPhone] : "N/A") || "N/A";
            otherPhoneDisplay.textContent = (colIdx.otherPhone !== -1 ? rowData[colIdx.otherPhone] : "N/A") || "N/A";
            studentEmailDisplay.textContent = (colIdx.studentEmail !== -1 ? rowData[colIdx.studentEmail] : "N/A") || "N/A";
            personalEmailDisplay.textContent = (colIdx.personalEmail !== -1 ? rowData[colIdx.personalEmail] : "N/A") || "N/A";

            const gender = colIdx.gender !== -1 ? String(rowData[colIdx.gender] || '').toLowerCase() : "";
            studentAvatar.textContent = getInitials(studentName);
            studentAvatar.style.backgroundColor = gender === 'female' ? '#ec4899' : gender === 'male' ? '#3b82f6' : '#6b7280';

            const daysOut = colIdx.daysOut !== -1 ? parseInt(rowData[colIdx.daysOut], 10) : null;
            daysOutDisplay.textContent = (daysOut !== null && !isNaN(daysOut)) ? daysOut : "--";
            daysOutStatBlock.className = 'flex-1 p-3 text-center rounded-lg';
            if (daysOut !== null && !isNaN(daysOut)) {
                if (daysOut >= 14) {
                    daysOutStatBlock.classList.add('bg-red-200', 'text-red-800');
                } else if (daysOut > 10) {
                    daysOutStatBlock.classList.add('bg-orange-200', 'text-orange-800');
                } else if (daysOut > 5) {
                    daysOutStatBlock.classList.add('bg-yellow-200', 'text-yellow-800');
                } else {
                    daysOutStatBlock.classList.add('bg-green-200', 'text-green-800');
                }
            } else {
                daysOutStatBlock.classList.add('bg-gray-200', 'text-gray-800');
            }

            let grade = colIdx.grade !== -1 ? rowData[colIdx.grade] : null;
            gradeStatBlock.className = 'flex-1 p-3 text-center rounded-lg';
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
