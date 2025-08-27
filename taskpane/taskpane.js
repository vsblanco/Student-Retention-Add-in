/*
 * This file contains the logic for the Task Pane UI.
 * It is loaded by taskpane.html.
 */

const CONSTANTS = {
    // UI Element IDs
    MAIN_CONTENT: "main-content",
    TAB_DETAILS: "tab-details",
    TAB_HISTORY: "tab-history",
    PANEL_DETAILS: "panel-details",
    PANEL_HISTORY: "panel-history",
    SUBMIT_COMMENT_BUTTON: "submit-comment-button",
    NEW_COMMENT_SECTION: "new-comment-section",
    ADD_COMMENT_BUTTON: "add-comment-button",
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
    COPY_STUDENT_ID: "copy-student-id",
    COPY_PRIMARY_PHONE: "copy-primary-phone",
    COPY_OTHER_PHONE: "copy-other-phone",
    COPY_STUDENT_EMAIL: "copy-student-email",
    COPY_PERSONAL_EMAIL: "copy-personal-email",
    USER_SELECTION_MODAL: "user-selection-modal",
    USER_SELECTION_DROPDOWN: "user-selection-dropdown",
    CONFIRM_USER_BUTTON: "confirm-user-button",
    CANCEL_USER_BUTTON: "cancel-user-button",
    TAG_PILLS_CONTAINER: "tag-pills-container",
    TAG_PLACEHOLDER: "tag-placeholder",
    ADD_TAG_BUTTON: "add-tag-button",
    TAG_DROPDOWN: "tag-dropdown",
    DATE_PICKER_MODAL: "date-picker-modal",

    // Settings Keys
    SETTINGS_KEY: "studentRetentionSettings",

    // Sheet and Column Names
    HISTORY_SHEET: "Student History",
    OUTREACH_HIGHLIGHT_TRIGGERS: ["will engage", "will submit", "will come","will complete"], // Phrases that trigger auto-highlight
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
        outreach: ["outreach"],
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
let sessionCommentUser = null; // Cache the user for the current session
let pendingOutreachAction = null; // Cache outreach data while waiting for user selection
let newCommentTags = []; // To store tags for the new comment
let selectedDate = null; // For the date picker
let smartLdaTag = null; // To track the automatically added LDA tag

const availableTags = [
    { name: 'Urgent', bg: 'bg-red-100', text: 'text-red-800' },
    { name: 'Note', bg: 'bg-gray-700', text: 'text-gray-100' },
    { name: 'LDA', bg: 'bg-orange-100', text: 'text-orange-800', requiresDate: true },
    { name: 'Contacted', bg: 'bg-yellow-100', text: 'text-yellow-800', hidden: true },
    { name: 'Outreach', bg: 'bg-blue-100', text: 'text-blue-800', hidden: true },
    { name: 'Quote', bg: 'bg-sky-100', text: 'text-sky-800', hidden: true }
];

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    loadUserSettings();
    
    // Add listeners for the new user selection modal buttons
    const confirmUserButton = document.getElementById(CONSTANTS.CONFIRM_USER_BUTTON);
    if (confirmUserButton) {
        confirmUserButton.addEventListener('click', handleUserSelection);
    }

    if (sessionCommentUser) {
        // If a user is already selected for the session, show the main content
        document.getElementById(CONSTANTS.MAIN_CONTENT).classList.remove('hidden');
        initializeAddIn();
    } else {
        // Otherwise, prompt for a user first.
        promptForUserAndSubmit();
    }
  }
});


/**
 * Initializes the main add-in functionality after a user has been selected.
 */
function initializeAddIn() {
    // By the time Office is ready, the DOM should be loaded as well.
    setupTabs();
    setupCopyHandlers(); // Set up the copy-to-clipboard functionality
    setupGradebookLinkHandler(); // Set up the gradebook link handler
    
    // Add event listener for the new comment button
    const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);
    if (submitButton) {
        submitButton.addEventListener("click", submitNewComment);
    }
    
    // Add event listener for the floating action button to toggle the new comment section
    const addCommentButton = document.getElementById(CONSTANTS.ADD_COMMENT_BUTTON);
    if (addCommentButton) {
        addCommentButton.addEventListener('click', () => {
            const commentSection = document.getElementById(CONSTANTS.NEW_COMMENT_SECTION);
            commentSection.classList.toggle('hidden');
            // If the section is now visible, focus the input
            if (!commentSection.classList.contains('hidden')) {
                document.getElementById(CONSTANTS.NEW_COMMENT_INPUT).focus();
            }
        });
    }
    
    const commentInput = document.getElementById(CONSTANTS.NEW_COMMENT_INPUT);
    if (commentInput) {
        // Add paste event listener to the comment input to automatically add the "Quote" tag
        commentInput.addEventListener('paste', () => {
            addTag('Quote');
        });
        // Add input event listener for smart LDA tag detection
        commentInput.addEventListener('input', handleCommentInputChange);
    }

    // Initialize the tagging UI
    setupTaggingUI();
    
    // Initialize the date picker
    setupDatePicker();

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
    
    // Ensure userProfile object and userList exist
    if (!settings.userProfile) {
        settings.userProfile = { userList: [] };
    }
    if (!settings.userProfile.userList) {
        settings.userProfile.userList = [];
    }
    if (!settings.taskpane) {
        settings.taskpane = { smartNavigation: true };
    }

    // Load user name from settings, fallback to Office context display name
    currentUserName = settings.userProfile.name || Office.context.displayName;

    // Add the current user to the list if not already present, then save
    if (currentUserName && !settings.userProfile.userList.includes(currentUserName)) {
        settings.userProfile.userList.push(currentUserName);
        saveUserSettings();
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

        // Add the new name to the user list if it's not already there
        if (!settings.userProfile.userList) {
            settings.userProfile.userList = [];
        }
        if (!settings.userProfile.userList.includes(currentUserName)) {
            settings.userProfile.userList.push(currentUserName);
        }
        
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
    
    tabDetails.addEventListener("click", () => switchTab(CONSTANTS.TAB_DETAILS));
    tabHistory.addEventListener("click", () => switchTab(CONSTANTS.TAB_HISTORY));
}

/**
 * Handles the logic for switching between tabs.
 * @param {string} tabId The ID of the tab to switch to.
 */
function switchTab(tabId) {
    const tabDetails = document.getElementById(CONSTANTS.TAB_DETAILS);
    const tabHistory = document.getElementById(CONSTANTS.TAB_HISTORY);
    const panelDetails = document.getElementById(CONSTANTS.PANEL_DETAILS);
    const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
    const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);
    const addCommentButton = document.getElementById(CONSTANTS.ADD_COMMENT_BUTTON);

    if (tabId === CONSTANTS.TAB_DETAILS) {
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        panelDetails.classList.remove("hidden");
        panelHistory.classList.add("hidden");
        if (addCommentButton) addCommentButton.classList.add('hidden');
    } else { // Switching to History
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        panelHistory.classList.remove("hidden");
        panelDetails.classList.add("hidden");
        if (addCommentButton) addCommentButton.classList.remove('hidden');
        
        if (submitButton) {
            submitButton.disabled = !currentStudentId;
        }

        if (currentStudentId) {
            displayStudentHistory(currentStudentId);
        } else {
            const historyContent = document.getElementById(CONSTANTS.HISTORY_CONTENT);
            historyContent.innerHTML = '<p class="text-gray-500">Select a student row to see their history.</p>';
        }
    }
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
 * This version uses UTC methods to prevent timezone-related date shifts.
 * @param {number} excelDate The Excel date serial number.
 * @returns {string} The formatted date string or "N/A".
 */
function formatExcelDate(excelDate) {
    if (isNaN(excelDate) || excelDate === null || excelDate === "") {
        return "N/A";
    }
    // Excel's epoch starts on 1900-01-01, but it incorrectly thinks 1900 is a leap year.
    // The JavaScript epoch starts on 1970-01-01. The difference is 25569 days.
    // The calculation correctly gives UTC milliseconds.
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    
    if (isNaN(date.getTime())) {
        return "N/A";
    }

    // Use UTC methods to avoid timezone shifts.
    const day = date.getUTCDate();
    const year = date.getUTCFullYear();
    const month = date.toLocaleString('default', { month: 'long', timeZone: 'UTC' });

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
            selectedRange.load(["rowIndex", "columnIndex", "address"]);
            
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load(["rowIndex", "values", "formulas"]);
            
            await context.sync();

            // --- Smart Navigation & Data Loading Logic ---
            
            const headers = usedRange.values[0];
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
                gradeBook: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook),
                outreach: findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.outreach)
            };

            // Smart Navigation: This runs on every selection change.
            if (settings.taskpane && settings.taskpane.smartNavigation !== false) {
                const selectedCol = selectedRange.columnIndex;
                if (selectedCol === colIdx.primaryPhone || selectedCol === colIdx.otherPhone || selectedCol === colIdx.outreach) {
                    switchTab(CONSTANTS.TAB_HISTORY);
                } else {
                    switchTab(CONSTANTS.TAB_DETAILS);
                }
            }

            // Data Loading: This part only runs if the row has changed.
            if (selectedRange.rowIndex === lastSelectedRow) {
                return; 
            }
            lastSelectedRow = selectedRange.rowIndex;

            if (selectedRange.rowIndex < usedRange.rowIndex) {
                currentStudentId = null;
                currentStudentName = null;
                return;
            }

            const rowDataIndex = lastSelectedRow - usedRange.rowIndex;
            if (rowDataIndex < 0 || rowDataIndex >= usedRange.values.length) {
                console.error("Selected row is outside the bounds of the used range data.");
                currentStudentId = null;
                currentStudentName = null;
                return;
            }
            const rowData = usedRange.values[rowDataIndex];
            const rowFormulas = usedRange.formulas[rowDataIndex];

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
        CONSTANTS.COPY_STUDENT_ID,
        CONSTANTS.COPY_PRIMARY_PHONE,
        CONSTANTS.COPY_OTHER_PHONE,
        CONSTANTS.COPY_STUDENT_EMAIL,
        CONSTANTS.COPY_PERSONAL_EMAIL
    ];

    contactInfoIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('click', () => {
                // For the student ID, the text is in a different child element
                const displayEl = id === CONSTANTS.COPY_STUDENT_ID
                    ? el.querySelector('#' + CONSTANTS.STUDENT_ID_DISPLAY)
                    : el.querySelector('.font-semibold');
                
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
 * Finds and highlights date-related text within a comment.
 * @param {string} text The comment text.
 * @returns {string} The comment text with the date part bolded.
 */
function highlightDateInText(text) {
    // Regex to find days of the week, tomorrow, or month-day combinations.
    // It's case-insensitive and handles optional suffixes like 'st', 'nd', 'rd', 'th'.
    const dateRegex = /\b(tomorrow|sunday|monday|tuesday|wednesday|thursday|friday|saturday|january|february|march|april|may|june|july|august|september|october|november|december)\b(\s+\d{1,2}(st|nd|rd|th)?)?/i;

    const match = text.match(dateRegex);

    if (match) {
        // Replace the first occurrence of the matched date string with a bolded version.
        return text.replace(dateRegex, (matchedText) => {
            return `<strong class="font-bold text-gray-900">${matchedText}</strong>`;
        });
    }

    return text; // Return original text if no date found
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
                        const comment = {
                            text: commentText,
                            tag: tagColIdx !== -1 ? row[tagColIdx] : null,
                            timestamp: timestampColIdx !== -1 ? row[timestampColIdx] : null,
                            createdBy: (createdByColIdx !== -1 && row[createdByColIdx]) ? row[createdByColIdx] : 'Unknown',
                            ldaDate: null // New property for priority sorting
                        };

                        // Check for future LDA tag
                        if (comment.tag) {
                            const tags = String(comment.tag).split(',').map(t => t.trim());
                            const ldaTag = tags.find(t => t.toLowerCase().startsWith('lda '));
                            if (ldaTag) {
                                const dateStr = ldaTag.substring(4);
                                const ldaDate = new Date(dateStr);
                                if (!isNaN(ldaDate.getTime())) {
                                    ldaDate.setHours(0, 0, 0, 0); // Normalize to start of day for comparison
                                    comment.ldaDate = ldaDate;
                                }
                            }
                        }
                        comments.push(comment);
                    }
                }
            }
            
            if (comments.length > 0) {
                // --- Sorting Logic ---
                const today = new Date();
                today.setHours(0, 0, 0, 0);

                const noteComments = comments.filter(c => c.tag && c.tag.toLowerCase().includes('note'));
                const priorityComments = comments.filter(c => c.ldaDate && c.ldaDate >= today && !noteComments.includes(c));
                const regularComments = comments.filter(c => !noteComments.includes(c) && !priorityComments.includes(c));

                // Sort note comments by timestamp (most recent first)
                noteComments.sort((a, b) => {
                    const dateA = parseDate(a.timestamp) || 0;
                    const dateB = parseDate(b.timestamp) || 0;
                    return dateB - dateA;
                });

                // Sort priority comments by LDA date, ascending (soonest first)
                priorityComments.sort((a, b) => a.ldaDate - b.ldaDate);

                // Sort regular comments by timestamp, descending (most recent first)
                regularComments.sort((a, b) => {
                    const dateA = parseDate(a.timestamp) || 0;
                    const dateB = parseDate(b.timestamp) || 0;
                    return dateB - dateA;
                });

                const sortedComments = [...noteComments, ...priorityComments, ...regularComments];
                // --- End Sorting Logic ---

                let html = '<ul class="space-y-4">';
                sortedComments.forEach(comment => {
                    const isNote = comment.tag && comment.tag.toLowerCase().includes('note');
                    const isPriority = !isNote && comment.ldaDate && comment.ldaDate >= today;
                    
                    let bgColor = 'bg-gray-100'; // Default
                    if (isNote) {
                        bgColor = 'bg-gray-200';
                    } else if (isPriority) {
                        bgColor = 'bg-orange-100';
                    }

                    let displayText = comment.text;
                    if (isPriority) {
                        displayText = highlightDateInText(comment.text);
                    }


                    html += `
                        <li class="p-3 ${bgColor} rounded-lg shadow-sm">
                            <p class="text-sm text-gray-800">${displayText}</p>`;
                    
                    html += `<div class="text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center">`;
                    html += `<div class="flex items-center gap-2">`; // Left side container
                    
                    // Tag rendering logic
                    if (comment.tag) {
                        const tags = String(comment.tag).split(',').map(t => t.trim());
                        tags.forEach(tagName => {
                            if (tagName && tagName.toLowerCase() !== 'comment') {
                                const tagPrefix = tagName.split(' ')[0];
                                const tagInfo = availableTags.find(t => t.name.toLowerCase() === tagPrefix.toLowerCase()) || 
                                                { bg: 'bg-blue-100', text: 'text-blue-800' }; // Default color
                                html += `<span class="px-2 py-0.5 font-semibold rounded-full ${tagInfo.bg} ${tagInfo.text}">${tagName}</span>`;
                            }
                        });
                    }
                    
                    html += `<span class="font-medium">${comment.createdBy}</span>`;
                    html += `</div>`; // End left side

                    let dateText;
                    if (comment.timestamp) {
                       dateText = comment.timestamp;
                       if (!isNaN(dateText) && dateText > 25569) {
                           const date = new Date((dateText - 25569) * 86400 * 1000);
                           dateText = date.toLocaleString('en-US', {
                               year: 'numeric',
                               month: 'numeric',
                               day: 'numeric',
                               hour: 'numeric',
                               minute: '2-digit'
                           });
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
 * Main entry point for submitting a comment. Decides whether to prompt for a user or proceed.
 */
async function submitNewComment() {
    if (sessionCommentUser) {
        await executeSubmitComment(sessionCommentUser);
    } else {
        promptForUserAndSubmit();
    }
}

/**
 * Displays a modal to ask the user who they are commenting as.
 */
function promptForUserAndSubmit() {
    const modal = document.getElementById(CONSTANTS.USER_SELECTION_MODAL);
    const dropdown = document.getElementById(CONSTANTS.USER_SELECTION_DROPDOWN);
    
    dropdown.innerHTML = ''; // Clear previous options

    const userList = (settings.userProfile && settings.userProfile.userList) || [currentUserName];
    userList.forEach(user => {
        const option = document.createElement('option');
        option.value = user;
        option.textContent = user;
        if (user === currentUserName) {
            option.selected = true; // Default to the current user
        }
        dropdown.appendChild(option);
    });

    modal.classList.remove('hidden');
}

/**
 * Handles the user's selection from the modal, caches it, and proceeds with submission.
 */
async function handleUserSelection() {
    const dropdown = document.getElementById(CONSTANTS.USER_SELECTION_DROPDOWN);
    const selectedUser = dropdown.value;
    sessionCommentUser = selectedUser; // Cache the selection for the session
    document.getElementById(CONSTANTS.USER_SELECTION_MODAL).classList.add('hidden');
    document.getElementById(CONSTANTS.MAIN_CONTENT).classList.remove('hidden');
    
    // If this was the initial prompt, we now initialize the rest of the add-in
    if (!pendingOutreachAction) {
        initializeAddIn();
    }

    // Check if there's a pending outreach action to complete
    if (pendingOutreachAction) {
        const { studentId, studentName, commentText, rowIndex } = pendingOutreachAction;
        
        // Add the comment
        await addOutreachComment(studentId, studentName, commentText, sessionCommentUser);

        // Check for highlight triggers
        const lowerCommentText = commentText.toLowerCase();
        if (CONSTANTS.OUTREACH_HIGHLIGHT_TRIGGERS.some(phrase => lowerCommentText.includes(phrase))) {
            console.log(`Highlight trigger phrase found for ${studentName}. Highlighting row ${rowIndex + 1}.`);
            await applyContactedHighlight(rowIndex);
        }

        // Clear the pending action
        pendingOutreachAction = null;
    } else {
        // If no pending action, it was triggered by the submit button
        await executeSubmitComment(sessionCommentUser);
    }
}

/**
 * Executes the logic to add a new comment to the "Student History" sheet.
 * @param {string} commentingUser The name of the user to attribute the comment to.
 */
async function executeSubmitComment(commentingUser) {
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
            
            const newRowData = new Array(headers.length).fill("");
            
            newRowData[idCol] = currentStudentId;
            if (studentCol !== -1) newRowData[studentCol] = currentStudentName;
            if (createdByCol !== -1) newRowData[createdByCol] = commentingUser;
            
            // Join the selected tags, or default to "Comment" if none are selected
            const tagsToSave = newCommentTags.length > 0 ? newCommentTags.join(', ') : "Comment";
            if (tagCol !== -1) newRowData[tagCol] = tagsToSave;

            if (timestampCol !== -1) {
                const now = new Date();
                newRowData[timestampCol] = (now.getTime() / 86400000) + 25569;
            }
            newRowData[commentCol] = commentText;

            const newRowRange = historySheet.getRangeByIndexes(newRowIndex, 0, 1, headers.length);
            newRowRange.values = [newRowData];
            
            historySheet.getUsedRange().getEntireColumn().format.autofitColumns();
            
            await context.sync();
        });

        commentInput.value = "";
        statusDisplay.textContent = "Comment added successfully!";
        
        // Reset tags and refresh history
        newCommentTags = [];
        smartLdaTag = null; 
        renderTagPills();
        populateTagDropdown();
        await displayStudentHistory(currentStudentId);

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

// --- START: Tagging UI Functions ---

function setupTaggingUI() {
    const addTagButton = document.getElementById(CONSTANTS.ADD_TAG_BUTTON);
    const tagDropdown = document.getElementById(CONSTANTS.TAG_DROPDOWN);

    addTagButton.addEventListener('click', (event) => {
        event.stopPropagation();
        populateTagDropdown();
        tagDropdown.classList.toggle('hidden');
    });

    // Hide dropdown if clicked outside
    document.addEventListener('click', (event) => {
        if (!tagDropdown.contains(event.target) && !addTagButton.contains(event.target)) {
            tagDropdown.classList.add('hidden');
        }
    });

    renderTagPills();
    populateTagDropdown();
}

function renderTagPills() {
    const container = document.getElementById(CONSTANTS.TAG_PILLS_CONTAINER);
    const placeholder = document.getElementById(CONSTANTS.TAG_PLACEHOLDER);

    // Add a guard clause to prevent errors if elements aren't found
    if (!container || !placeholder) {
        return;
    }

    // Remove existing pills, but not the placeholder
    container.querySelectorAll('.tag-pill').forEach(pill => pill.remove());

    if (newCommentTags.length === 0) {
        placeholder.classList.remove('hidden');
    } else {
        placeholder.classList.add('hidden');
        newCommentTags.forEach(tagName => {
            const tagPrefix = tagName.split(' ')[0];
            const tagInfo = availableTags.find(t => t.name === tagPrefix) || { bg: 'bg-gray-200', text: 'text-gray-800' };
            const pill = document.createElement('span');
            // Add a class to identify these as generated pills
            pill.className = `tag-pill px-2 py-1 text-xs font-semibold rounded-full flex items-center gap-1 ${tagInfo.bg} ${tagInfo.text}`;
            pill.textContent = tagName;

            const removeButton = document.createElement('button');
            removeButton.type = 'button';
            removeButton.innerHTML = '&times;';
            removeButton.className = 'font-bold';
            removeButton.onclick = () => removeTag(tagName);

            pill.appendChild(removeButton);
            container.appendChild(pill);
        });
    }
}

function populateTagDropdown() {
    const dropdown = document.getElementById(CONSTANTS.TAG_DROPDOWN);
    if (!dropdown) return; // Guard clause
    dropdown.innerHTML = '';

    const tagsToShow = availableTags.filter(tag => !tag.hidden && !newCommentTags.some(t => t.startsWith(tag.name)));

    if (tagsToShow.length === 0) {
        const noTagsItem = document.createElement('span');
        noTagsItem.className = 'block px-4 py-2 text-sm text-gray-500';
        noTagsItem.textContent = 'No more tags';
        dropdown.appendChild(noTagsItem);
        return;
    }

    tagsToShow.forEach(tag => {
        const item = document.createElement('a');
        item.href = '#';
        item.className = 'block px-3 py-2 text-sm text-gray-700 hover:bg-gray-100';
        
        const pill = document.createElement('span');
        pill.className = `px-2 py-0.5 text-xs font-semibold rounded-full ${tag.bg} ${tag.text}`;
        pill.textContent = tag.name;
        item.appendChild(pill);

        item.onclick = (e) => {
            e.preventDefault();
            if (tag.requiresDate) {
                promptForLdaDate();
            } else {
                addTag(tag.name);
            }
        };
        dropdown.appendChild(item);
    });
}

function addTag(tagName) {
    if (!newCommentTags.some(t => t.startsWith(tagName.split(' ')[0]))) {
        newCommentTags.push(tagName);
        renderTagPills();
        populateTagDropdown();
        document.getElementById(CONSTANTS.TAG_DROPDOWN).classList.add('hidden');
    }
}

function removeTag(tagName) {
    newCommentTags = newCommentTags.filter(t => t !== tagName);
    if (tagName === smartLdaTag) {
        smartLdaTag = null;
    }
    renderTagPills();
    populateTagDropdown();
}


// --- END: Tagging UI Functions ---

// --- START: Date Picker Functions ---

function setupDatePicker() {
    let currentDate = new Date();
    
    document.getElementById('prev-month').addEventListener('click', () => {
        currentDate.setMonth(currentDate.getMonth() - 1);
        renderCalendar(currentDate);
    });

    document.getElementById('next-month').addEventListener('click', () => {
        currentDate.setMonth(currentDate.getMonth() + 1);
        renderCalendar(currentDate);
    });
    
    document.getElementById('cancel-date').addEventListener('click', () => {
        document.getElementById(CONSTANTS.DATE_PICKER_MODAL).classList.add('hidden');
    });
    
    document.getElementById('confirm-date').addEventListener('click', () => {
        if (selectedDate) {
            const formattedDate = `${selectedDate.getMonth() + 1}/${selectedDate.getDate()}/${String(selectedDate.getFullYear()).slice(-2)}`;
            addTag(`LDA ${formattedDate}`);
            document.getElementById(CONSTANTS.DATE_PICKER_MODAL).classList.add('hidden');
        }
    });
}

function promptForLdaDate() {
    selectedDate = null; // Reset selection
    document.getElementById('confirm-date').disabled = true;
    renderCalendar(new Date());
    document.getElementById(CONSTANTS.DATE_PICKER_MODAL).classList.remove('hidden');
    document.getElementById(CONSTANTS.TAG_DROPDOWN).classList.add('hidden');
}

function renderCalendar(date) {
    const grid = document.getElementById('calendar-grid');
    const header = document.getElementById('month-year');
    
    grid.innerHTML = ''; // Clear old calendar
    const month = date.getMonth();
    const year = date.getFullYear();

    header.textContent = `${date.toLocaleString('default', { month: 'long' })} ${year}`;

    // Day headers
    ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach(day => {
        const dayEl = document.createElement('div');
        dayEl.className = 'font-bold text-gray-500';
        dayEl.textContent = day;
        grid.appendChild(dayEl);
    });

    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    // Blank spaces for the first week
    for (let i = 0; i < firstDay; i++) {
        grid.appendChild(document.createElement('div'));
    }

    // Date cells
    for (let i = 1; i <= daysInMonth; i++) {
        const dayEl = document.createElement('button');
        dayEl.textContent = i;
        dayEl.className = 'p-1 rounded-full hover:bg-blue-200';
        
        const dayDate = new Date(year, month, i);
        
        if (selectedDate && dayDate.getTime() === selectedDate.getTime()) {
            dayEl.classList.add('bg-blue-500', 'text-white');
        }
        
        dayEl.onclick = () => {
            selectedDate = dayDate;
            document.getElementById('confirm-date').disabled = false;
            renderCalendar(date); // Re-render to show selection
        };
        grid.appendChild(dayEl);
    }
}


// --- END: Date Picker Functions ---

// --- START: Smart Tag Functions ---

/**
 * Handles the input event on the comment box to check for smart tags.
 * @param {Event} event The input event.
 */
function handleCommentInputChange(event) {
    const text = event.target.value.toLowerCase();
    const triggerWords = ['engage', 'submit'];

    const hasTrigger = triggerWords.some(word => text.includes(word));
    const date = hasTrigger ? parseDateFromText(text) : null;

    // If a date is found, add or update the smart LDA tag
    if (date) {
        const formattedDate = `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
        const newLdaTag = `LDA ${formattedDate}`;

        // If a smart tag already exists but is different, remove the old one
        if (smartLdaTag && smartLdaTag !== newLdaTag) {
            removeTag(smartLdaTag);
        }

        // If no smart tag exists (or it was just removed), add the new one
        if (!smartLdaTag) {
            // Also check if a manual LDA tag was added, don't add if so
            if (!newCommentTags.some(tag => tag.startsWith('LDA'))) {
                 addTag(newLdaTag);
                 smartLdaTag = newLdaTag;
            }
        }
    } else {
        // If no date is found, but we had a smart tag, remove it
        if (smartLdaTag) {
            removeTag(smartLdaTag);
        }
    }
}

/**
 * Parses a date from a string of text, understanding relative terms.
 * @param {string} text The text to parse.
 * @returns {Date|null} A Date object or null if no date is found.
 */
function parseDateFromText(text) {
    const now = new Date();
    text = text.toLowerCase();

    // Check for "tomorrow"
    if (text.includes('tomorrow')) {
        const tomorrow = new Date();
        tomorrow.setDate(now.getDate() + 1);
        return tomorrow;
    }

    // Check for days of the week
    const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    for (let i = 0; i < days.length; i++) {
        if (text.includes(days[i])) {
            const targetDay = i;
            const today = now.getDay();
            let daysToAdd = targetDay - today;
            if (daysToAdd <= 0) {
                daysToAdd += 7; // Get the next occurrence
            }
            const nextDate = new Date();
            nextDate.setDate(now.getDate() + daysToAdd);
            return nextDate;
        }
    }
    
    // Check for "Month Day" format (e.g., "August 20th")
    const monthDayRegex = /(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\s+(\d{1,2})/;
    const monthDayMatch = text.match(monthDayRegex);
    if (monthDayMatch) {
        const months = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
                         january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11 };
        const month = months[monthDayMatch[1]];
        const day = parseInt(monthDayMatch[2], 10);
        let year = now.getFullYear();
        const date = new Date(year, month, day);
        if (date < now) { // If the date is in the past, assume next year
            date.setFullYear(year + 1);
        }
        return date;
    }

    return null;
}


// --- END: Smart Tag Functions ---


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
 * Applies a yellow highlight to a specific row to mark it as contacted.
 * This is a non-toggle version of the function in actions.js.
 * @param {number} rowIndex The zero-based index of the row to highlight.
 */
async function applyContactedHighlight(rowIndex) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const headerRange = sheet.getRange("1:1").getUsedRange(true);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0];
            const lowerCaseHeaders = headers.map(header => String(header || '').toLowerCase());
            const studentNameColIndex = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.name);
            const outreachColIndex = findColumnIndex(lowerCaseHeaders, CONSTANTS.COLUMN_MAPPINGS.outreach);

            if (studentNameColIndex === -1 || outreachColIndex === -1) {
                console.error("Could not find 'StudentName' and/or 'Outreach' columns for highlighting.");
                return;
            }

            const startCol = Math.min(studentNameColIndex, outreachColIndex);
            const endCol = Math.max(studentNameColIndex, outreachColIndex);
            const colCount = endCol - startCol + 1;

            const highlightRange = sheet.getRangeByIndexes(rowIndex, startCol, 1, colCount);
            highlightRange.format.fill.color = "yellow";
            await context.sync();
        });
    } catch (error) {
        errorHandler(error);
    }
}


/**
 * Event handler for when the worksheet changes. Now handles bulk pastes.
 * @param {Excel.WorksheetChangedEventArgs} eventArgs
 */
async function onWorksheetChanged(eventArgs) {
    if (!sessionCommentUser) {
        console.log("Outreach change ignored: No session user selected.");
        return;
    }

    await Excel.run(async (context) => {
        if (eventArgs.source !== Excel.EventSource.local) return;
        if (eventArgs.changeType !== "CellEdited" && eventArgs.changeType !== "RangeEdited") return;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const changedRange = sheet.getRange(eventArgs.address);
        
        const headerRange = sheet.getRange("1:1").getUsedRange(true);
        headerRange.load("values, columnCount");
        changedRange.load("address, rowIndex, columnIndex, rowCount, columnCount, values, valuesBefore");
        await context.sync();
        
        const headers = (headerRange.values[0] || []).map(h => String(h || '').toLowerCase());
        const outreachColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.outreach);
        
        if (outreachColIndex === -1) return;

        if (changedRange.rowIndex > 0 && 
            changedRange.columnIndex <= outreachColIndex && 
            (changedRange.columnIndex + changedRange.columnCount - 1) >= outreachColIndex) {

            const outreachColumnOffset = outreachColIndex - changedRange.columnIndex;
            
            const studentInfoRange = sheet.getRangeByIndexes(
                changedRange.rowIndex, 0, 
                changedRange.rowCount, headerRange.columnCount
            );
            studentInfoRange.load("values");
            await context.sync();
            const allRowValues = studentInfoRange.values;

            const studentIdColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.id);
            const studentNameColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.name);

            if (studentIdColIndex === -1 || studentNameColIndex === -1) return;

            for (let i = 0; i < changedRange.rowCount; i++) {
                const newValue = (changedRange.values[i] && changedRange.values[i][outreachColumnOffset]) ? 
                                 String(changedRange.values[i][outreachColumnOffset] || "").trim() : "";
                
                const oldValue = (changedRange.valuesBefore && changedRange.valuesBefore[i] && changedRange.valuesBefore[i][outreachColumnOffset]) ?
                                 String(changedRange.valuesBefore[i][outreachColumnOffset] || "").trim() : "";

                if (newValue !== "" && newValue.toLowerCase() !== oldValue.toLowerCase()) {
                    const studentId = allRowValues[i][studentIdColIndex];
                    const studentName = allRowValues[i][studentNameColIndex];
                    const rowIndex = changedRange.rowIndex + i;

                    if (studentId && studentName) {
                        await addOutreachComment(studentId, studentName, newValue, sessionCommentUser);
                        const lowerNewValue = newValue.toLowerCase();
                        if (CONSTANTS.OUTREACH_HIGHLIGHT_TRIGGERS.some(phrase => lowerNewValue.includes(phrase))) {
                            console.log(`Highlight trigger phrase found for ${studentName}. Highlighting row ${rowIndex + 1}.`);
                            await applyContactedHighlight(rowIndex);
                        }
                    }
                }
            }
        }
    }).catch(errorHandler);
}


/**
 * Adds or updates a comment in the "Student History" sheet based on an outreach entry.
 * If an outreach comment for the student exists for the current day, it updates it.
 * Otherwise, it adds a new comment.
 * @param {string|number} studentId The student's ID.
 * @param {string} studentName The student's name.
 * @param {string} commentText The new comment text from the Outreach column.
 * @param {string} commentingUser The user making the comment.
 */
async function addOutreachComment(studentId, studentName, commentText, commentingUser) {
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

            if (idCol === -1 || commentCol === -1 || timestampCol === -1 || tagCol === -1) {
                console.log("Student History sheet is missing required columns (StudentNumber, Comment, Timestamp, Tag).");
                return;
            }

            let todaysCommentRowIndex = -1;
            const today = new Date();
            today.setHours(0, 0, 0, 0);

            // Search backwards for the most recent entry for this student today
            for (let i = historyData.length - 1; i > 0; i--) {
                const row = historyData[i];
                if (row[idCol] && String(row[idCol]) === String(studentId)) {
                    const tags = String(row[tagCol] || '').toLowerCase();
                    if (tags.includes('outreach')) {
                        const commentDate = parseDate(row[timestampCol]);
                        if (commentDate) {
                            commentDate.setHours(0, 0, 0, 0);
                            if (commentDate.getTime() === today.getTime()) {
                                todaysCommentRowIndex = historyRange.rowIndex + i;
                                break; // Found today's outreach comment
                            }
                        }
                    }
                }
            }

            const now = new Date();
            const excelNow = jsDateToExcelDate(now);

            if (todaysCommentRowIndex !== -1) {
                // Update existing comment
                console.log(`Updating existing outreach comment for ${studentName} at row ${todaysCommentRowIndex + 1}`);
                const commentCell = historySheet.getCell(todaysCommentRowIndex, commentCol);
                const timestampCell = historySheet.getCell(todaysCommentRowIndex, timestampCol);
                
                commentCell.values = [[commentText]];
                timestampCell.values = [[excelNow]];
                timestampCell.numberFormat = [["M/D/YYYY h:mm AM/PM"]];

            } else {
                // Add new comment
                console.log(`Adding new outreach comment for ${studentName}`);
                const newRowData = new Array(historyHeaders.length).fill("");
                newRowData[idCol] = studentId;
                if (studentCol !== -1) newRowData[studentCol] = studentName;
                if (createdByCol !== -1) newRowData[createdByCol] = commentingUser;
                
                let tagsToSave = "Outreach";
                const lowerCommentText = commentText.toLowerCase();

                // Check for highlight/contacted triggers
                if (CONSTANTS.OUTREACH_HIGHLIGHT_TRIGGERS.some(phrase => lowerCommentText.includes(phrase))) {
                    tagsToSave += ', Contacted';
                }

                const date = parseDateFromText(commentText);
                if (date) {
                    const formattedDate = `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
                    const ldaTag = `LDA ${formattedDate}`;
                    tagsToSave += `, ${ldaTag}`;
                }

                newRowData[tagCol] = tagsToSave;
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

            // Refresh the history pane if it's open for the current student
            if (studentId && String(studentId) === String(currentStudentId)) {
                const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
                if (panelHistory && !panelHistory.classList.contains("hidden")) {
                    console.log(`Refreshing history for student ${studentName} after outreach update.`);
                    setTimeout(() => displayStudentHistory(currentStudentId), 100);
                }
            }

        } catch (error) {
            errorHandler(error);
        }
    });
}
// --- END: Code moved from commands.js ---
