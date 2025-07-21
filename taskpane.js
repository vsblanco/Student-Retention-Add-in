/*
 * This file contains the logic for the Task Pane UI.
 * It is loaded by taskpane.html.
 */

let lastSelectedRow = -1; // Variable to track the last selected row index
let currentStudentId = null; // Variable to store the currently selected student's ID
let currentStudentName = null; // Variable to store the currently selected student's name

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // By the time Office is ready, the DOM should be loaded as well,
    // because we moved the script to the end of the body.
    setupTabs();
    
    // Add event listener for the new comment button
    const submitButton = document.getElementById("submit-comment-button");
    if (submitButton) {
        submitButton.addEventListener("click", submitNewComment);
    }

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to register selection change handler: " + result.error.message);
      } else {
        console.log("Selection change handler registered successfully.");
      }
    });
    
    // Run initial check
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
    const submitButton = document.getElementById("submit-comment-button");

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
        
        // Enable or disable submit button based on whether a student is selected
        if (submitButton) {
            submitButton.disabled = !currentStudentId;
        }

        // Fetch and display history when the tab is clicked
        if (currentStudentId) {
            displayStudentHistory(currentStudentId);
        } else {
            const historyContent = document.getElementById("history-content");
            historyContent.innerHTML = '<p class="text-gray-500">Select a student row to see their history.</p>';
        }
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
            selectedRange.load(["rowIndex", "address"]);
            await context.sync();

            if (selectedRange.rowIndex === lastSelectedRow) {
                return; 
            }
            lastSelectedRow = selectedRange.rowIndex;

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            const usedRange = sheet.getUsedRange();
            usedRange.load(["rowIndex", "values"]);
            await context.sync();

            if (selectedRange.rowIndex < usedRange.rowIndex) {
                currentStudentId = null;
                currentStudentName = null;
                return;
            }

            const headers = usedRange.values[0];
            const rowDataIndex = lastSelectedRow - usedRange.rowIndex;

            if (rowDataIndex < 0 || rowDataIndex >= usedRange.values.length) {
                console.error("Selected row is outside the bounds of the used range data.");
                currentStudentId = null;
                currentStudentName = null;
                return;
            }
            const rowData = usedRange.values[rowDataIndex];
            
            const lowerCaseHeaders = headers.map(header => String(header || '').toLowerCase());

            const columnMappings = {
                name: ["studentname", "student name"],
                id: ["student id", "studentnumber", "student identifier"],
                gender: ["gender"],
                daysOut: ["days out", "daysout"],
                grade: ["grade", "course grade"],
                status: ["status"],
                lastLda: ["last lda", "lda"],
                primaryPhone: ["primary phone", "phone"],
                otherPhone: ["other phone", "cell phone", "cell","otherphone"],
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
            
            const studentId = colIdx.id !== -1 ? rowData[colIdx.id] : null;
            currentStudentId = studentId;
            currentStudentName = studentName;

            const submitButton = document.getElementById('submit-comment-button');
            if (submitButton) {
                submitButton.disabled = !currentStudentId;
            }

            const panelHistory = document.getElementById("panel-history");
            if (!panelHistory.classList.contains("hidden")) {
                if (currentStudentId) {
                    displayStudentHistory(currentStudentId);
                } else {
                    const historyContent = document.getElementById("history-content");
                    historyContent.innerHTML = '<p class="text-gray-500">Could not find Student ID in the selected row.</p>';
                }
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
 * Fetches and displays the comment history for a given student ID from the "Student History" sheet.
 */
async function displayStudentHistory(studentId) {
    const historyContent = document.getElementById("history-content");
    historyContent.innerHTML = '<p class="text-gray-500">Loading history...</p>';

    try {
        await Excel.run(async (context) => {
            const historySheet = context.workbook.worksheets.getItem("Student History");
            const historyRange = historySheet.getUsedRange();
            historyRange.load("values");
            await context.sync();

            const historyData = historyRange.values;
            const historyHeaders = historyData[0].map(header => String(header || '').toLowerCase());
            
            const idColIdx = findColumnIndex(historyHeaders, ["student id", "student identifier"]);
            const commentColIdx = historyHeaders.indexOf("comment");
            const tagColIdx = historyHeaders.indexOf("tag");
            const timestampColIdx = historyHeaders.indexOf("timestamp");
            const createdByColIdx = historyHeaders.indexOf("created by");

            if (idColIdx === -1 || commentColIdx === -1) {
                historyContent.innerHTML = '<p class="text-red-500 font-semibold">Error: "Student History" sheet must contain "Student Identifier" and "Comment" columns.</p>';
                return;
            }

            const comments = [];
            for (let i = 1; i < historyData.length; i++) {
                const row = historyData[i];
                if (row[idColIdx] && String(row[idColIdx]) === String(studentId)) {
                    const commentText = row[commentColIdx];
                    if (commentText && String(commentText).trim() !== "") {
                        comments.push({
                            text: commentText,
                            tag: tagColIdx !== -1 ? row[tagColIdx] : null,
                            timestamp: timestampColIdx !== -1 ? row[timestampColIdx] : null,
                            createdBy: (createdByColIdx !== -1 && row[createdByColIdx]) ? row[createdByColIdx] : 'Unknown'
                        });
                    }
                }
            }
            
            if (comments.length > 0) {
                let html = '<ul class="space-y-4">';
                comments.reverse().forEach(comment => {
                    html += `
                        <li class="p-3 bg-gray-100 rounded-lg shadow-sm">
                            <p class="text-sm text-gray-800">${comment.text}</p>`;
                    
                    html += `<div class="text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center">`;
                    html += `<div>`; // Left side container
                    if (comment.tag) {
                        html += `<span class="px-2 py-0.5 font-semibold text-blue-800 bg-blue-100 rounded-full mr-2">${comment.tag}</span>`;
                    }
                    html += `<span class="font-medium">${comment.createdBy}</span>`;
                    html += `</div>`; // End left side

                    let dateText;
                    if (comment.timestamp) {
                       dateText = comment.timestamp;
                       if (!isNaN(dateText) && dateText > 25569) {
                           const date = new Date((dateText - 25569) * 86400 * 1000);
                           dateText = date.toLocaleString();
                       }
                    } else {
                        dateText = 'Unknown Time';
                    }
                    html += `<span>${dateText}</span>`;
                    
                    html += `</div>`;
                    html += `</li>`;
                });
                html += '</ul>';
                historyContent.innerHTML = html;
            } else {
                historyContent.innerHTML = '<p class="text-gray-500">No history found for this student.</p>';
            }
        });
    } catch (error) {
        if (error.code === 'ItemNotFound') {
            historyContent.innerHTML = '<p class="text-orange-500 font-semibold">A worksheet named "Student History" was not found.</p>';
        } else {
            historyContent.innerHTML = `<p class="text-red-500 font-semibold">An error occurred: ${error.message}</p>`;
            console.error("Error in displayStudentHistory: " + error);
        }
    }
}

/**
 * Submits a new comment to the "Student History" sheet.
 */
async function submitNewComment() {
    const commentInput = document.getElementById("new-comment-input");
    const statusDisplay = document.getElementById("comment-status");
    const commentText = commentInput.value.trim();

    if (!currentStudentId) {
        statusDisplay.textContent = "Please select a student first.";
        return;
    }
    if (commentText === "") {
        statusDisplay.textContent = "Comment cannot be empty.";
        return;
    }

    statusDisplay.textContent = "Submitting...";

    try {
        await Excel.run(async (context) => {
            const historySheet = context.workbook.worksheets.getItem("Student History");
            const historyRange = historySheet.getUsedRange();
            historyRange.load(["rowCount", "values"]);
            await context.sync();

            const newRowIndex = historyRange.rowCount;
            const headers = historyRange.values[0].map(h => String(h || '').toLowerCase());

            const idCol = findColumnIndex(headers, ["student id", "student identifier"]);
            const studentCol = headers.indexOf("student");
            const createdByCol = headers.indexOf("created by");
            const tagCol = headers.indexOf("tag");
            const timestampCol = headers.indexOf("timestamp");
            const commentCol = headers.indexOf("comment");

            if (idCol === -1 || commentCol === -1) {
                statusDisplay.textContent = "History sheet is missing required columns.";
                return;
            }
            
            // Create an array with the correct number of columns
            const newRowData = new Array(headers.length).fill("");
            
            // Populate the data based on found column indices
            newRowData[idCol] = currentStudentId;
            if (studentCol !== -1) newRowData[studentCol] = currentStudentName;
            if (createdByCol !== -1) newRowData[createdByCol] = "Victor Blanco"; // As per your name
            if (tagCol !== -1) newRowData[tagCol] = "Comment";
            if (timestampCol !== -1) {
                const now = new Date();
                // Convert JS date to Excel date number
                newRowData[timestampCol] = (now.getTime() / 86400000) + 25569;
            }
            newRowData[commentCol] = commentText;

            const newRowRange = historySheet.getRangeByIndexes(newRowIndex, 0, 1, headers.length);
            newRowRange.values = [newRowData];
            
            // Auto-fit columns for better readability
            historySheet.getUsedRange().getEntireColumn().format.autofitColumns();
            
            await context.sync();
        });

        commentInput.value = "";
        statusDisplay.textContent = "Comment added successfully!";
        // Refresh the history view
        await displayStudentHistory(currentStudentId);

        // Clear the status message after a few seconds
        setTimeout(() => {
            statusDisplay.textContent = "";
        }, 3000);

    } catch (error) {
        if (error.code === 'ItemNotFound') {
            statusDisplay.textContent = 'Error: "Student History" sheet not found.';
        } else {
            statusDisplay.textContent = `Error: ${error.message}`;
        }
        console.error("Error in submitNewComment: " + error);
    }
}
