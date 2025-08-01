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

    // Settings Keys
    SETTINGS_KEY: "studentRetentionSettings",

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
        outreach: ["outreach"], // Added from commands.js
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
let currentUserName = "Unknown User"; // Variable to store the current user's name
let settings = {}; // To store all add-in settings
let welcomeDialog = null;

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    loadUserSettings();

    // By the time Office is ready, the DOM should be loaded as well.
    setupTabs();
    setupCopyHandlers(); // Set up the copy-to-clipboard functionality
    setupGradebookLinkHandler(); // Set up the gradebook link handler
    
    // Add event listener for the new comment button
    const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);
    if (submitButton) {
        submitButton.addEventListener("click", submitNewComment);
    }

    // Add event handler for selection changes to update the task pane
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to register selection change handler: " + result.error.message);
      } else {
        console.log("Selection change handler registered successfully.");
      }
    });
    
    // Add event handler for worksheet changes to catch outreach updates
    Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.onChanged.add(onWorksheetChanged);
      await context.sync();
      console.log("Worksheet onChanged event handler registered from taskpane.");
    }).catch(errorHandler);

    // Run initial check
    onSelectionChange();
  }
});

/**
 * Loads all user settings from the document.
 */
function loadUserSettings() {
    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEY);
    if (settingsString) {
        try {
            settings = JSON.parse(settingsString);
        } catch (e) {
            console.error("Error parsing settings:", e);
            settings = {};
        }
    }
    
    // Ensure userProfile object exists
    if (!settings.userProfile) {
        settings.userProfile = {};
    }

    // Load user name
    if (settings.userProfile.name) {
        currentUserName = settings.userProfile.name;
    } else {
        currentUserName = Office.context.displayName;
    }
    
    // Show welcome message if needed
    if (!settings.userProfile.hasSeenWelcomeMessage) {
        showWelcomeDialogIfNeeded();
    }
}

/**
 * Saves all user settings to the document.
 */
function saveUserSettings() {
    Office.context.document.settings.set(CONSTANTS.SETTINGS_KEY, JSON.stringify(settings));
    Office.context.document.settings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to save settings: " + asyncResult.error.message);
        } else {
            console.log("Settings saved successfully.");
        }
    });
}


/**
 * Checks if the user has seen the welcome message and shows a dialog if not.
 */
function showWelcomeDialogIfNeeded() {
    if (!settings.userProfile.hasSeenWelcomeMessage) {
        Office.context.ui.displayDialogAsync(
            'https://vsblanco.github.io/Student-Retention-Add-in/welcome-dialog.html',
            { height: 40, width: 30, displayInIframe: true },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Welcome dialog failed to open: " + asyncResult.error.message);
                    return;
                }
                welcomeDialog = asyncResult.value;
                welcomeDialog.addEventHandler(Office.EventType.DialogMessageReceived, processWelcomeDialogMessage);
            }
        );
    }
}

/**
 * Handles messages from the welcome dialog.
 * @param {Office.DialogMessageReceivedEventArgs} arg
 */
function processWelcomeDialogMessage(arg) {
    const message = JSON.parse(arg.message);
    if (message.type === 'userName' && message.name) {
        currentUserName = message.name;
        settings.userProfile.name = currentUserName;
        settings.userProfile.hasSeenWelcomeMessage = true;
        saveUserSettings();
        
        if (welcomeDialog) {
            welcomeDialog.close();
            welcomeDialog = null;
        }
    }
}


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
                if (gradePercent >= 70) {
                    gradeStatBlock.classList.add('bg-green-200', 'text-green-800');
                } else if (gradePercent >= 60) {
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

                // Check if the color is already in our cache
                if (assignedColorMap[assignedTo]) {
                    const cellColor = assignedColorMap[assignedTo];
                    assignedToBadge.style.backgroundColor = cellColor;
                    const r = parseInt(cellColor.substr(1, 2), 16);
                    const g = parseInt(cellColor.substr(3, 2), 16);
                    const b = parseInt(cellColor.substr(5, 2), 16);
                    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
                    assignedToBadge.style.color = brightness > 125 ? 'black' : 'white';
                } else if (assignedTo) {
                    // If not in cache, get the color for this specific cell and cache it
                    const assignedCell = sheet.getCell(lastSelectedRow, colIdx.assigned);
                    assignedCell.load("format/fill/color");
                    await context.sync();

                    if (assignedCell.format && assignedCell.format.fill) {
                        const cellColor = assignedCell.format.fill.color;
                        if (cellColor && cellColor !== '#ffffff' && cellColor !== '#000000') {
                            assignedColorMap[assignedTo] = cellColor;
                            console.log(`[DEBUG] Caching new color for '${assignedTo}': ${cellColor}`);
                            
                            assignedToBadge.style.backgroundColor = cellColor;
                            const r = parseInt(cellColor.substr(1, 2), 16);
                            const g = parseInt(cellColor.substr(3, 2), 16);
                            const b = parseInt(cellColor.substr(5, 2), 16);
                            const brightness = (r * 299 + g * 587 + b * 114) / 1000;
                            assignedToBadge.style.color = brightness > 125 ? 'black' : 'white';
                        } else {
                            assignedToBadge.style.backgroundColor = '#e5e7eb';
                            assignedToBadge.style.color = '#1f2937';
                        }
                    }
                } else {
                    assignedToBadge.style.backgroundColor = '#e5e7eb';
                    assignedToBadge.style.color = '#1f2937';
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
            if (createdByCol !== -1) newRowData[createdByCol] = currentUserName;
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

// --- START: Code moved from commands.js ---

/**
 * Generic error handler for Excel.run calls.
 * @param {any} error The error object.
 */
function errorHandler(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}

/**
 * Parses a date value from various possible formats (Date object, string, Excel serial number).
 * @param {*} dateValue The value to parse.
 * @returns {Date|null} A valid Date object or null.
 */
function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
        return dateValue;
    }
    if (typeof dateValue === 'number') {
        if (dateValue > 25569) { // Corresponds to 1970-01-01
            return new Date((dateValue - 25569) * 86400 * 1000);
        }
    }
    if (typeof dateValue === 'string') {
        const parsed = new Date(dateValue);
        if (!isNaN(parsed.getTime())) {
            return parsed;
        }
    }
    return null;
}

/**
 * Converts a JavaScript Date object to an Excel serial date number.
 * @param {Date} date The JavaScript Date object.
 * @returns {number} The Excel serial date number.
 */
function jsDateToExcelDate(date) {
    return (date.getTime() / 86400000) + 25569;
}

/**
 * Event handler for when the worksheet changes. Now handles bulk pastes.
 * @param {Excel.WorksheetChangedEventArgs} eventArgs
 */
async function onWorksheetChanged(eventArgs) {
    await Excel.run(async (context) => {
        // Only trigger for changes made by the local user
        if (eventArgs.source !== Excel.EventSource.local) {
            console.log("Change from remote source ignored.");
            return;
        }

        if (eventArgs.changeType !== "CellEdited" && eventArgs.changeType !== "RangeEdited") {
            return;
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const changedRange = sheet.getRange(eventArgs.address);
        
        const headerRange = sheet.getRange("1:1").getUsedRange(true);
        headerRange.load("values, columnCount");
        changedRange.load("address, rowIndex, columnIndex, rowCount, columnCount, values, valuesBefore");

        await context.sync();
        
        const headers = (headerRange.values[0] || []).map(h => String(h || '').toLowerCase());
        const outreachColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.outreach);
        
        if (outreachColIndex === -1) {
            return;
        }

        // Check if the change is in the data area and intersects with the outreach column
        if (changedRange.rowIndex > 0 && 
            changedRange.columnIndex <= outreachColIndex && 
            (changedRange.columnIndex + changedRange.columnCount - 1) >= outreachColIndex) {

            const outreachColumnOffset = outreachColIndex - changedRange.columnIndex;
            
            // Load all necessary data for the affected rows in one go
            const studentInfoRange = sheet.getRangeByIndexes(
                changedRange.rowIndex, 0, 
                changedRange.rowCount, headerRange.columnCount
            );
            studentInfoRange.load("values");
            await context.sync();
            const allRowValues = studentInfoRange.values;

            const studentIdColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.id);
            const studentNameColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.name);

            if (studentIdColIndex === -1 || studentNameColIndex === -1) {
                console.log("Required columns (StudentNumber, StudentName) not found.");
                return;
            }

            // Loop through each row in the changed range
            for (let i = 0; i < changedRange.rowCount; i++) {
                const newValue = (changedRange.values[i] && changedRange.values[i][outreachColumnOffset]) ? 
                                 String(changedRange.values[i][outreachColumnOffset] || "").trim() : "";
                
                // FIX: Add a check for changedRange.valuesBefore before accessing its properties
                const oldValue = (changedRange.valuesBefore && changedRange.valuesBefore[i] && changedRange.valuesBefore[i][outreachColumnOffset]) ?
                                 String(changedRange.valuesBefore[i][outreachColumnOffset] || "").trim() : "";

                if (newValue !== "" && newValue.toLowerCase() !== oldValue.toLowerCase()) {
                    const studentId = allRowValues[i][studentIdColIndex];
                    const studentName = allRowValues[i][studentNameColIndex];

                    if (studentId && studentName) {
                        console.log(`Processing pasted value for ${studentName}: "${newValue}"`);
                        await addOutreachComment(studentId, studentName, newValue);
                    }
                }
            }
        }
    }).catch(errorHandler);
}


/**
 * Adds or updates a comment in the "Student History" sheet based on an outreach entry.
 * @param {string|number} studentId The student's ID.
 * @param {string} studentName The student's name.
 * @param {string} commentText The new comment text from the Outreach column.
 */
async function addOutreachComment(studentId, studentName, commentText) {
    await Excel.run(async (context) => {
        try {
            const historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
            const historyRange = historySheet.getUsedRange(true);
            historyRange.load("values, rowIndex, rowCount, address");
            await context.sync();

            const historyData = historyRange.values;
            const historyHeaders = historyData[0].map(h => String(h || '').toLowerCase());

            const idCol = findColumnIndex(historyHeaders, CONSTANTS.COLUMN_MAPPINGS.id);
            const commentCol = findColumnIndex(historyHeaders, ["comment"]);
            const timestampCol = findColumnIndex(historyHeaders, ["timestamp"]);
            const tagCol = findColumnIndex(historyHeaders, ["tag"]);
            const createdByCol = findColumnIndex(historyHeaders, ["created by"]);
            const studentCol = findColumnIndex(historyHeaders, ["student"]);

            if (idCol === -1 || commentCol === -1 || timestampCol === -1) {
                console.log("Student History sheet is missing required columns (StudentNumber, Comment, Timestamp).");
                return;
            }

            let lastCommentRowIndex = -1;
            let lastCommentTimestamp = 0;

            for (let i = historyData.length - 1; i > 0; i--) {
                const row = historyData[i];
                if (row[idCol] && String(row[idCol]) === String(studentId)) {
                    lastCommentRowIndex = historyRange.rowIndex + i;
                    lastCommentTimestamp = row[timestampCol];
                    break;
                }
            }

            const now = new Date();
            const excelNow = jsDateToExcelDate(now);
            const oneMinuteInMillis = 60 * 1000;
            let updateExisting = false;

            if (lastCommentRowIndex !== -1) {
                const lastCommentDate = parseDate(lastCommentTimestamp);
                if (lastCommentDate && (now.getTime() - lastCommentDate.getTime()) < oneMinuteInMillis) {
                    updateExisting = true;
                }
            }

            if (updateExisting) {
                console.log(`Updating existing comment for ${studentName} at row ${lastCommentRowIndex + 1}`);
                const commentCell = historySheet.getCell(lastCommentRowIndex, commentCol);
                commentCell.values = [[commentText]];
                const timestampCell = historySheet.getCell(lastCommentRowIndex, timestampCol);
                timestampCell.values = [[excelNow]];
                timestampCell.numberFormat = [["M/D/YYYY h:mm AM/PM"]];

            } else {
                console.log(`Adding new comment for ${studentName}`);
                const newRowData = new Array(historyHeaders.length).fill("");
                newRowData[idCol] = studentId;
                if (studentCol !== -1) newRowData[studentCol] = studentName;
                if (createdByCol !== -1) newRowData[createdByCol] = currentUserName;
                if (tagCol !== -1) newRowData[tagCol] = "Outreach";
                newRowData[timestampCol] = excelNow;
                newRowData[commentCol] = commentText;
                
                const newRowIndex = historyRange.rowIndex + historyRange.rowCount;
                const newRowRange = historySheet.getRangeByIndexes(newRowIndex, 0, 1, historyHeaders.length);
                newRowRange.values = [newRowData];
                
                const newTimestampCell = historySheet.getCell(newRowIndex, timestampCol);
                newTimestampCell.numberFormat = [["M/D/YYYY h:mm AM/PM"]];
            }
            
            historySheet.getUsedRange().format.autofitColumns();
            await context.sync();

            // Check if the updated student is the one currently displayed in the task pane
            // and if the history tab is active. If so, refresh the history.
            if (studentId && String(studentId) === String(currentStudentId)) {
                const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
                if (panelHistory && !panelHistory.classList.contains("hidden")) {
                    console.log(`Refreshing history for student ${studentName} after outreach update.`);
                    // Use a brief timeout to allow the UI to process the sheet change before refreshing.
                    setTimeout(() => displayStudentHistory(currentStudentId), 100);
                }
            }

        } catch (error) {
            errorHandler(error);
        }
    });
}
// --- END: Code moved from commands.js ---
