/*
 * This file contains the logic for the Task Pane UI.
 * It is loaded by taskpane.html.
 */

const CONSTANTS = {
    // UI Element IDs
    TAB_DETAILS: "tab-details",
    TAB_HISTORY: "tab-history",
    PANEL_DETAILS: "panel-details",
    PANEL_HISTORY: "panel-history",
    SUBMIT_COMMENT_BUTTON: "submit-comment-button",
    NEW_COMMENT_INPUT: "new-comment-input",
    COMMENT_STATUS: "comment-status",
    HISTORY_CONTENT: "history-content",
    STUDENT_AVATAR: "student-avatar",
    STUDENT_NAME_DISPLAY: "student-name-display",
    ASSIGNED_TO_BADGE: "assigned-to-badge",
    STUDENT_ID_DISPLAY: "student-id-display",
    LAST_LDA_DISPLAY: "last-lda-display",
    DAYS_OUT_DISPLAY: "days-out-display",
    DAYS_OUT_STAT_BLOCK: "days-out-stat-block",
    GRADE_DISPLAY: "grade-display",
    GRADE_STAT_BLOCK: "grade-stat-block",
    PRIMARY_PHONE_DISPLAY: "primary-phone-display",
    OTHER_PHONE_DISPLAY: "other-phone-display",
    STUDENT_EMAIL_DISPLAY: "student-email-display",
    PERSONAL_EMAIL_DISPLAY: "personal-email-display",
    COPY_PRIMARY_PHONE: "copy-primary-phone",
    COPY_OTHER_PHONE: "copy-other-phone",
    COPY_STUDENT_EMAIL: "copy-student-email",
    COPY_PERSONAL_EMAIL: "copy-personal-email",

    // Sheet and Column Names
    HISTORY_SHEET: "Student History",
    COLUMN_MAPPINGS: {
        name: ["studentname", "student name"],
        id: ["student id", "studentnumber", "student identifier"],
        gender: ["gender"],
        daysOut: ["days out", "daysout"],
        grade: ["grade", "course grade"],
        status: ["status"],
        assigned: ["assigned"],
        lastLda: ["last lda", "lda"],
        primaryPhone: ["primary phone", "phone"],
        otherPhone: ["other phone", "cell phone", "cell", "otherphone"],
        studentEmail: ["student email", "school email", "email"],
        personalEmail: ["personal email", "otheremail"],
        gradeBook: ["grade book", "gradebook"],
        comment: "comment",
        tag: "tag",
        timestamp: "timestamp",
        createdBy: "created by",
        student: "student"
    }
};


let lastSelectedRow = -1; // Variable to track the last selected row index
let currentStudentId = null; // Variable to store the currently selected student's ID
let currentStudentName = null; // Variable to store the currently selected student's name
let currentGradebookLink = null; // Variable to store the gradebook link
let assignedColorMap = {}; // To cache colors for assigned people

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // By the time Office is ready, the DOM should be loaded as well.
    setupTabs();
    setupCopyHandlers(); // Set up the copy-to-clipboard functionality
    setupGradebookLinkHandler(); // Set up the gradebook link handler
    
    // Add event listener for the new comment button
    const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);
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
    
    // Run initial check and cache the colors
    onSelectionChange();
    cacheAssignedColors();
  }
});

/**
 * Sets up the event listeners for the tabbed interface.
 */
function setupTabs() {
    const tabDetails = document.getElementById(CONSTANTS.TAB_DETAILS);
    const tabHistory = document.getElementById(CONSTANTS.TAB_HISTORY);
    const panelDetails = document.getElementById(CONSTANTS.PANEL_DETAILS);
    const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
    const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);

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
            const historyContent = document.getElementById(CONSTANTS.HISTORY_CONTENT);
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
 * Converts an Excel date serial number to a formatted string (e.g., "June 1st, 2025").
 * @param {number} excelDate The Excel date serial number.
 * @returns {string} The formatted date string or "N/A".
 */
function formatExcelDate(excelDate) {
    if (isNaN(excelDate) || excelDate === null || excelDate === "") {
        return "N/A";
    }
    // Excel's epoch starts on 1900-01-01, but it incorrectly thinks 1900 is a leap year.
    // The JavaScript epoch starts on 1970-01-01. The difference is 25569 days.
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    
    if (isNaN(date.getTime())) {
        return "N/A";
    }

    const day = date.getDate();
    const year = date.getFullYear();
    const month = date.toLocaleString('default', { month: 'long' });

    let daySuffix = 'th';
    if (day === 1 || day === 21 || day === 31) {
        daySuffix = 'st';
    } else if (day === 2 || day === 22) {
        daySuffix = 'nd';
    } else if (day === 3 || day === 23) {
        daySuffix = 'rd';
    }

    return `${month} ${day}${daySuffix}, ${year}`;
}

/**
 * Scans the "Assigned" column and caches the fill color for each unique person.
 */
async function cacheAssignedColors() {
    console.log("[DEBUG] Starting cacheAssignedColors...");
    try {
        await Excel.run(async (context) => {
            console.log("[DEBUG] Excel.run for cacheAssignedColors started.");
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            
            // First, load only the values to find the column index efficiently.
            usedRange.load("values, rowCount");
            await context.sync();

            const headers = usedRange.values[0].map(header => String(header || '').toLowerCase());
            const assignedColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.assigned);

            if (assignedColIdx === -1) {
                console.log("[DEBUG] 'Assigned' column not found.");
                return;
            }
            console.log(`[DEBUG] 'Assigned' column found at index: ${assignedColIdx}`);

            // Now, get a specific range for the data in that column (excluding the header).
            // This is a more targeted and reliable way to get individual cell formats.
            const assignedColumnDataRange = sheet.getRangeByIndexes(1, assignedColIdx, usedRange.rowCount - 1, 1);
            assignedColumnDataRange.load("values, format/fill/color");
            await context.sync();
            console.log(`[DEBUG] Loaded data for specific range: ${assignedColumnDataRange.address}`);

            const newColorMap = {};
            const values = assignedColumnDataRange.values;
            const colors = assignedColumnDataRange.format.fill.color;

            if (!colors) {
                console.log("[DEBUG] No fill color information available for the 'Assigned' column data range.");
                return;
            }

            // Iterate through each row of the column data.
            for (let i = 0; i < values.length; i++) {
                const name = values[i][0];

                // If the name is valid and we haven't cached a color for it yet...
                if (name && !newColorMap[name]) {
                    const cellColor = colors[i][0];
                    // Cache the color if it's not white or black (default colors).
                    if (cellColor && cellColor !== '#ffffff' && cellColor !== '#000000') {
                        newColorMap[name] = cellColor;
                        console.log(`[DEBUG] Caching color for '${name}': ${cellColor}`);
                    }
                }
            }
            
            assignedColorMap = newColorMap;
            console.log("[DEBUG] Final assigned colors cached:", assignedColorMap);
        });
    } catch (error) {
        console.error("Error caching assigned colors: " + error.message);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info for cacheAssignedColors: " + JSON.stringify(error.debugInfo));
        }
    }
}


/**
 * Handles the document selection change event.
 */
async function onSelectionChange() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load(["rowIndex", "address"]);
            
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            // Load formulas to detect hyperlinks
            usedRange.load(["rowIndex", "values", "formulas"]);
            
            await context.sync();

            if (selectedRange.rowIndex === lastSelectedRow) {
                return; 
            }
            lastSelectedRow = selectedRange.rowIndex;

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
            const rowFormulas = usedRange.formulas[rowDataIndex];
            
            const lowerCaseHeaders = headers.map(header => String(header || '').toLowerCase());

            const colIdx = {
                name: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.name),
                id: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.id),
                gender: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.gender),
                daysOut: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut),
                grade: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.grade),
                status: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.status),
                assigned: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.assigned),
                lastLda: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.lastLda),
                primaryPhone: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.primaryPhone),
                otherPhone: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.otherPhone),
                studentEmail: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.studentEmail),
                personalEmail: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.personalEmail),
                gradeBook: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook)
            };

            const studentAvatar = document.getElementById(CONSTANTS.STUDENT_AVATAR);
            const studentNameDisplay = document.getElementById(CONSTANTS.STUDENT_NAME_DISPLAY);
            const assignedToBadge = document.getElementById(CONSTANTS.ASSIGNED_TO_BADGE);
            const studentIdDisplay = document.getElementById(CONSTANTS.STUDENT_ID_DISPLAY);
            const lastLdaDisplay = document.getElementById(CONSTANTS.LAST_LDA_DISPLAY);
            const daysOutDisplay = document.getElementById(CONSTANTS.DAYS_OUT_DISPLAY);
            const daysOutStatBlock = document.getElementById(CONSTANTS.DAYS_OUT_STAT_BLOCK);
            const gradeDisplay = document.getElementById(CONSTANTS.GRADE_DISPLAY);
            const gradeStatBlock = document.getElementById(CONSTANTS.GRADE_STAT_BLOCK);
            const primaryPhoneDisplay = document.getElementById(CONSTANTS.PRIMARY_PHONE_DISPLAY);
            const otherPhoneDisplay = document.getElementById(CONSTANTS.OTHER_PHONE_DISPLAY);
            const studentEmailDisplay = document.getElementById(CONSTANTS.STUDENT_EMAIL_DISPLAY);
            const personalEmailDisplay = document.getElementById(CONSTANTS.PERSONAL_EMAIL_DISPLAY);

            const studentName = colIdx.name !== -1 ? rowData[colIdx.name] : "N/A";
            
            studentNameDisplay.textContent = studentName || "N/A";
            studentIdDisplay.textContent = (colIdx.id !== -1 ? rowData[colIdx.id] : "N/A") || "N/A";
            
            const lastLdaValue = colIdx.lastLda !== -1 ? rowData[colIdx.lastLda] : null;
            lastLdaDisplay.textContent = formatExcelDate(lastLdaValue);

            primaryPhoneDisplay.textContent = (colIdx.primaryPhone !== -1 ? rowData[colIdx.primaryPhone] : "N/A") || "N/A";
            otherPhoneDisplay.textContent = (colIdx.otherPhone !== -1 ? rowData[colIdx.otherPhone] : "N/A") || "N/A";
            studentEmailDisplay.textContent = (colIdx.studentEmail !== -1 ? rowData[colIdx.studentEmail] : "N/A") || "N/A";
            personalEmailDisplay.textContent = (colIdx.personalEmail !== -1 ? rowData[colIdx.personalEmail] : "N/A") || "N/A";

            const gender = colIdx.gender !== -1 ? String(rowData[colIdx.gender] || '').toLowerCase() : "";
            studentAvatar.textContent = getInitials(studentName);
            studentAvatar.style.backgroundColor = gender === 'female' ? '#ec4899' : gender === 'male' ? '#3b82f6' : '#6b7280';

            const daysOut = colIdx.daysOut !== -1 ? parseInt(rowData[colIdx.daysOut], 10) : null;
            daysOutDisplay.textContent = (daysOut !== null && !isNaN(daysOut)) ? daysOut : "--";
            daysOutStatBlock.className = 'flex-1 p-3 text-center rounded-lg bg-gray-200 text-gray-800';
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
            gradeStatBlock.className = 'flex-1 p-3 text-center rounded-lg transition-colors duration-150';
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

            // Check for hyperlink formula first
            const gradeBookFormula = (colIdx.gradeBook !== -1 && rowFormulas) ? rowFormulas[colIdx.gradeBook] : null;
            const hyperlinkRegex = /=HYPERLINK\("([^"]+)"/i;
            const match = gradeBookFormula ? String(gradeBookFormula).match(hyperlinkRegex) : null;

            if (match && match[1]) {
                currentGradebookLink = match[1];
                gradeStatBlock.classList.add('cursor-pointer', 'hover:bg-gray-300');
                gradeStatBlock.title = 'Click to open Grade Book';
            } else {
                // Fallback for plain text URLs
                const gradeBookValue = colIdx.gradeBook !== -1 ? rowData[colIdx.gradeBook] : null;
                if (gradeBookValue && (String(gradeBookValue).startsWith('http://') || String(gradeBookValue).startsWith('https://'))) {
                    currentGradebookLink = gradeBookValue;
                    gradeStatBlock.classList.add('cursor-pointer', 'hover:bg-gray-300');
                    gradeStatBlock.title = 'Click to open Grade Book';
                } else {
                    currentGradebookLink = null;
                    gradeStatBlock.classList.remove('cursor-pointer', 'hover:bg-gray-300');
                    gradeStatBlock.title = '';
                }
            }

            // Handle Assigned To badge
            if (colIdx.assigned !== -1) {
                const assignedTo = rowData[colIdx.assigned];
                assignedToBadge.textContent = assignedTo || "Unassigned";

                const cellColor = assignedColorMap[assignedTo]; // Use the cache

                if (cellColor) {
                    assignedToBadge.style.backgroundColor = cellColor;
                    // Simple brightness check to determine text color
                    const r = parseInt(cellColor.substr(1, 2), 16);
                    const g = parseInt(cellColor.substr(3, 2), 16);
                    const b = parseInt(cellColor.substr(5, 2), 16);
                    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
                    assignedToBadge.style.color = brightness > 125 ? 'black' : 'white';
                } else {
                    assignedToBadge.style.backgroundColor = '#e5e7eb'; // Tailwind gray-200
                    assignedToBadge.style.color = '#1f2937'; // Tailwind gray-800
                }
            } else {
                assignedToBadge.textContent = "Unassigned";
                assignedToBadge.style.backgroundColor = '#e5e7eb';
                assignedToBadge.style.color = '#1f2937';
            }

            const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);
            if (submitButton) {
                submitButton.disabled = !currentStudentId;
            }

            const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
            if (!panelHistory.classList.contains("hidden")) {
                if (currentStudentId) {
                    displayStudentHistory(currentStudentId);
                } else {
                    const historyContent = document.getElementById(CONSTANTS.HISTORY_CONTENT);
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
 * Copies text to the clipboard and provides user feedback.
 * @param {string} text The text to copy.
 * @param {HTMLElement} triggerElement The element that triggered the copy action.
 */
function copyToClipboard(text, triggerElement) {
    if (!text || text === "N/A") {
        return; // Don't copy placeholder text
    }

    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "absolute";
    textArea.style.left = "-9999px";
    document.body.appendChild(textArea);
    textArea.select();
    try {
        document.execCommand('copy');
        const feedbackEl = triggerElement.querySelector('.copy-feedback');
        if (feedbackEl) {
            // Hide all other feedback elements first
            document.querySelectorAll('.copy-feedback').forEach(el => el.classList.add('hidden'));
            feedbackEl.classList.remove('hidden');
            setTimeout(() => {
                feedbackEl.classList.add('hidden');
            }, 2000);
        }
    } catch (err) {
        console.error('Failed to copy text: ', err);
    }
    document.body.removeChild(textArea);
}

/**
 * Sets up click event listeners for contact information to enable copy-to-clipboard.
 */
function setupCopyHandlers() {
    const contactInfoIds = [
        CONSTANTS.COPY_PRIMARY_PHONE,
        CONSTANTS.COPY_OTHER_PHONE,
        CONSTANTS.COPY_STUDENT_EMAIL,
        CONSTANTS.COPY_PERSONAL_EMAIL
    ];

    contactInfoIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('click', () => {
                const displayEl = el.querySelector('.font-semibold');
                if (displayEl) {
                    copyToClipboard(displayEl.textContent, el);
                }
            });
        }
    });
}

/**
 * Sets up the click event listener for the gradebook link.
 */
function setupGradebookLinkHandler() {
    const gradeStatBlock = document.getElementById(CONSTANTS.GRADE_STAT_BLOCK);
    if (gradeStatBlock) {
        gradeStatBlock.addEventListener('click', () => {
            if (currentGradebookLink) {
                window.open(currentGradebookLink, '_blank');
            }
        });
    }
}

/**
 * Fetches and displays the comment history for a given student ID from the "Student History" sheet.
 */
async function displayStudentHistory(studentId) {
    const historyContent = document.getElementById(CONSTANTS.HISTORY_CONTENT);
    historyContent.innerHTML = '<p class="text-gray-500">Loading history...</p>';

    try {
        await Excel.run(async (context) => {
            const historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
            const historyRange = historySheet.getUsedRange();
            historyRange.load("values");
            await context.sync();

            const historyData = historyRange.values;
            const historyHeaders = historyData[0].map(header => String(header || '').toLowerCase());
            
            const idColIdx = findColumnIndex(historyHeaders, CONSTANTS.COLUMN_MAPPINGS.id);
            const commentColIdx = historyHeaders.indexOf(CONSTANTS.COLUMN_MAPPINGS.comment);
            const tagColIdx = historyHeaders.indexOf(CONSTANTS.COLUMN_MAPPINGS.tag);
            const timestampColIdx = historyHeaders.indexOf(CONSTANTS.COLUMN_MAPPINGS.timestamp);
            const createdByColIdx = historyHeaders.indexOf(CONSTANTS.COLUMN_MAPPINGS.createdBy);

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
            historyContent.innerHTML = `<p class="text-orange-500 font-semibold">A worksheet named "${CONSTANTS.HISTORY_SHEET}" was not found.</p>`;
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
    const commentInput = document.getElementById(CONSTANTS.NEW_COMMENT_INPUT);
    const statusDisplay = document.getElementById(CONSTANTS.COMMENT_STATUS);
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
            const historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
            const historyRange = historySheet.getUsedRange();
            historyRange.load(["rowCount", "values"]);
            await context.sync();

            const newRowIndex = historyRange.rowCount;
            const headers = historyRange.values[0].map(h => String(h || '').toLowerCase());

            const idCol = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.id);
            const studentCol = headers.indexOf(CONSTANTS.COLUMN_MAPPINGS.student);
            const createdByCol = headers.indexOf(CONSTANTS.COLUMN_MAPPINGS.createdBy);
            const tagCol = headers.indexOf(CONSTANTS.COLUMN_MAPPINGS.tag);
            const timestampCol = headers.indexOf(CONSTANTS.COLUMN_MAPPINGS.timestamp);
            const commentCol = headers.indexOf(CONSTANTS.COLUMN_MAPPINGS.comment);

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
            statusDisplay.textContent = `Error: "${CONSTANTS.HISTORY_SHEET}" sheet not found.`;
        } else {
            statusDisplay.textContent = `Error: ${error.message}`;
        }
        console.error("Error in submitNewComment: " + error);
    }
}
