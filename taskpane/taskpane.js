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
    STUDENT_ID_DISPLAY_COPY: "student-id-display-copy",
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
    USER_BUTTONS_CONTAINER: "user-buttons-container",
    TAG_PILLS_CONTAINER: "tag-pills-container",
    ADD_TAG_BUTTON: "add-tag-button",
    TAG_DROPDOWN: "tag-dropdown",
    DATE_PICKER_MODAL: "date-picker-modal",
    EDIT_COMMENT_MODAL: "edit-comment-modal",
    UPDATE_COMMENT_BUTTON: "update-comment-button",
    CANCEL_EDIT_BUTTON: "cancel-edit-button",
    EDIT_COMMENT_INPUT: "edit-comment-input",
    EDIT_TAG_PILLS_CONTAINER: "edit-tag-pills-container",
    EDIT_ADD_TAG_BUTTON: "edit-add-tag-button",
    EDIT_TAG_DROPDOWN: "edit-tag-dropdown",
    EDIT_COMMENT_STATUS: "edit-comment-status",
    COMMENT_CONTEXT_MENU: "comment-context-menu",
    EDIT_COMMENT_BTN: "edit-comment-btn",
    DELETE_COMMENT_BTN: "delete-comment-btn",
    COPY_COMMENT_BTN: "copy-comment-btn",
    DNC_TYPE_MODAL: "dnc-type-modal",
    DNC_OPTIONS_CONTAINER: "dnc-options-container",
    CANCEL_DNC_BUTTON: "cancel-dnc-button",
    DELETE_CONFIRM_MODAL: "delete-confirm-modal",
    CONFIRM_DELETE_BUTTON: "confirm-delete-button",
    CANCEL_DELETE_BUTTON: "cancel-delete-button",
    SEARCH_HISTORY_BUTTON: "search-history-button",
    SEARCH_CONTAINER: "search-container",
    SEARCH_INPUT: "search-input",
    CLEAR_SEARCH_BUTTON: "clear-search-button",
    TAG_FILTER_CONTAINER: "tag-filter-container",
    FILTER_ADD_TAG_BUTTON: "filter-add-tag-button",
    FILTER_TAG_DROPDOWN: "filter-tag-dropdown",


    // Settings Keys
    SETTINGS_KEY: "studentRetentionSettings",

    // Sheet and Column Names
    HISTORY_SHEET: "Student History",
    OUTREACH_HIGHLIGHT_TRIGGERS: ["will engage", "will submit", "will come","will complete","will work","hanged up"],
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


let lastSelectedRow = -1;
let currentStudentId = null;
let currentStudentName = null;
let currentGradebookLink = null;
let assignedColorMap = {};
let currentUserName = "Unknown User";
let settings = {};
let welcomeDialog = null;
let sessionCommentUser = null;
let pendingOutreachAction = null;
let newCommentTags = [];
let selectedDate = null;
let smartLdaTag = null;
let allComments = [];
let activeCommentRowIndex = null;
let editCommentTags = [];
let activeTaggingMode = 'new';
let activeFilterTags = [];

const availableTags = [
    { name: 'Urgent', bg: 'bg-red-100', text: 'text-red-800', description: 'For high-priority items needing immediate attention.' },
    { name: 'Note', bg: 'bg-gray-700', text: 'text-gray-100', description: 'A general-purpose note for internal records.' },
    { name: 'DNC', bg: 'bg-red-700', text: 'text-black', requiresPopup: true, description: 'Do Not Contact. Opens options for Phone, Email, or All.' },
    { name: 'LDA', bg: 'bg-orange-100', text: 'text-orange-800', requiresDate: true, description: 'Last Date of Attendance. Adds a future follow-up date.' },
    { name: 'Contacted', bg: 'bg-yellow-100', text: 'text-yellow-800', description: 'Marks that contact has been made with the student.' },
    { name: 'Outreach', bg: 'bg-blue-100', text: 'text-blue-800', hidden: true },
    { name: 'Quote', bg: 'bg-sky-100', text: 'text-sky-800', hidden: true }
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    loadUserSettings();
    
    if (sessionCommentUser) {
        document.getElementById(CONSTANTS.MAIN_CONTENT).classList.remove('hidden');
        initializeAddIn();
    } else {
        promptForUserAndSubmit();
    }
  }
});


/**
 * Initializes the main add-in functionality after a user has been selected.
 */
function initializeAddIn() {
    setupTabs();
    setupCopyHandlers();
    setupGradebookLinkHandler();
    
    const submitButton = document.getElementById(CONSTANTS.SUBMIT_COMMENT_BUTTON);
    if (submitButton) {
        submitButton.addEventListener("click", submitNewComment);
    }
    
    const addCommentButton = document.getElementById(CONSTANTS.ADD_COMMENT_BUTTON);
    if (addCommentButton) {
        addCommentButton.addEventListener('click', () => {
            const commentSection = document.getElementById(CONSTANTS.NEW_COMMENT_SECTION);
            commentSection.classList.toggle('hidden');
            if (!commentSection.classList.contains('hidden')) {
                document.getElementById(CONSTANTS.NEW_COMMENT_INPUT).focus();
            }
        });
    }
    
    const commentInput = document.getElementById(CONSTANTS.NEW_COMMENT_INPUT);
    if (commentInput) {
        commentInput.addEventListener('paste', () => addTag('Quote', 'new'));
        commentInput.addEventListener('input', handleCommentInputChange);
    }

    setupTaggingUI('new');
    setupTaggingUI('edit');
    setupTaggingUI('filter');
    setupDatePicker();
    setupCommentEditing();
    setupDncModal();
    setupDeleteConfirmation();
    setupSearch(); 
    renderTagFilters();

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to register selection change handler: " + result.error.message);
      } else {
        console.log("Selection change handler registered successfully.");
      }
    });
    
    Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.onChanged.add(onWorksheetChanged);
      await context.sync();
      console.log("Worksheet onChanged event handler registered from taskpane.");
    }).catch(errorHandler);

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
    
    if (!settings.userProfile) settings.userProfile = { userList: [] };
    if (!settings.userProfile.userList) settings.userProfile.userList = [];
    if (!settings.taskpane) settings.taskpane = { smartNavigation: true };

    currentUserName = settings.userProfile.name || Office.context.displayName;

    if (currentUserName && !settings.userProfile.userList.includes(currentUserName)) {
        settings.userProfile.userList.push(currentUserName);
        saveUserSettings();
    }
    
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

        if (!settings.userProfile.userList) settings.userProfile.userList = [];
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

    if (tabId === CONSTANTS.TAB_DETAILS) {
        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        panelDetails.classList.remove("hidden");

        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        panelHistory.classList.add("hidden");
    } else { // Switching to History
        tabHistory.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600";
        panelHistory.classList.remove("hidden");
        panelHistory.classList.add("flex"); // Ensure flex is on for sticky header

        tabDetails.className = "whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300";
        panelDetails.classList.add("hidden");
        
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
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    
    if (isNaN(date.getTime())) {
        return "N/A";
    }

    const day = date.getUTCDate();
    const year = date.getUTCFullYear();
    const month = date.toLocaleString('default', { month: 'long', timeZone: 'UTC' });

    let daySuffix = 'th';
    if (day === 1 || day === 21 || day === 31) daySuffix = 'st';
    else if (day === 2 || day === 22) daySuffix = 'nd';
    else if (day === 3 || day === 23) daySuffix = 'rd';

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

            if (settings.taskpane && settings.taskpane.smartNavigation !== false) {
                const selectedCol = selectedRange.columnIndex;
                if (selectedCol === colIdx.primaryPhone || selectedCol === colIdx.otherPhone || selectedCol === colIdx.outreach) {
                    switchTab(CONSTANTS.TAB_HISTORY);
                } else {
                    switchTab(CONSTANTS.TAB_DETAILS);
                }
            }

            if (selectedRange.rowIndex === lastSelectedRow) return; 
            lastSelectedRow = selectedRange.rowIndex;

            if (selectedRange.rowIndex < usedRange.rowIndex) {
                currentStudentId = null;
                currentStudentName = null;
                return;
            }

            const rowDataIndex = lastSelectedRow - usedRange.rowIndex;
            if (rowDataIndex < 0 || rowDataIndex >= usedRange.values.length) {
                currentStudentId = null;
                currentStudentName = null;
                return;
            }
            const rowData = usedRange.values[rowDataIndex];
            const rowFormulas = usedRange.formulas[rowDataIndex];

            const studentAvatar = document.getElementById(CONSTANTS.STUDENT_AVATAR);
            const studentNameDisplay = document.getElementById(CONSTANTS.STUDENT_NAME_DISPLAY);
            const assignedToBadge = document.getElementById(CONSTANTS.ASSIGNED_TO_BADGE);
            const studentIdDisplayCopy = document.getElementById(CONSTANTS.STUDENT_ID_DISPLAY_COPY);
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
            studentIdDisplayCopy.textContent = (colIdx.id !== -1 ? rowData[colIdx.id] : "N/A") || "N/A";
            
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
                if (daysOut >= 14) daysOutStatBlock.classList.add('bg-red-200', 'text-red-800');
                else if (daysOut > 10) daysOutStatBlock.classList.add('bg-orange-200', 'text-orange-800');
                else if (daysOut > 5) daysOutStatBlock.classList.add('bg-yellow-200', 'text-yellow-800');
                else daysOutStatBlock.classList.add('bg-green-200', 'text-green-800');
            } else {
                daysOutStatBlock.classList.add('bg-gray-200', 'text-gray-800');
            }

            let grade = colIdx.grade !== -1 ? rowData[colIdx.grade] : null;
            gradeStatBlock.className = 'flex-1 p-3 text-center rounded-lg transition-colors duration-150';
            if (grade !== null && !isNaN(grade)) {
                const gradePercent = grade > 1 ? grade : grade * 100;
                gradeDisplay.textContent = `${Math.round(gradePercent)}%`;
                if (gradePercent >= 70) gradeStatBlock.classList.add('bg-green-200', 'text-green-800');
                else if (gradePercent >= 60) gradeStatBlock.classList.add('bg-yellow-200', 'text-yellow-800');
                else gradeStatBlock.classList.add('bg-red-200', 'text-red-800');
            } else {
                gradeDisplay.textContent = 'N/A';
                gradeStatBlock.classList.add('bg-gray-200', 'text-gray-800');
            }
            
            const studentId = colIdx.id !== -1 ? rowData[colIdx.id] : null;
            currentStudentId = studentId;
            currentStudentName = studentName;

            const gradeBookFormula = (colIdx.gradeBook !== -1 && rowFormulas) ? rowFormulas[colIdx.gradeBook] : null;
            const hyperlinkRegex = /=HYPERLINK\("([^"]+)"/i;
            const match = gradeBookFormula ? String(gradeBookFormula).match(hyperlinkRegex) : null;

            if (match && match[1]) {
                currentGradebookLink = match[1];
                gradeStatBlock.classList.add('cursor-pointer', 'hover:bg-gray-300');
                gradeStatBlock.title = 'Click to open Grade Book';
            } else {
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

            if (colIdx.assigned !== -1) {
                const assignedTo = rowData[colIdx.assigned];
                assignedToBadge.textContent = assignedTo || "Unassigned";

                if (assignedColorMap[assignedTo]) {
                    const cellColor = assignedColorMap[assignedTo];
                    assignedToBadge.style.backgroundColor = cellColor;
                    const r = parseInt(cellColor.substr(1, 2), 16), g = parseInt(cellColor.substr(3, 2), 16), b = parseInt(cellColor.substr(5, 2), 16);
                    assignedToBadge.style.color = ((r * 299 + g * 587 + b * 114) / 1000) > 125 ? 'black' : 'white';
                } else if (assignedTo) {
                    const assignedCell = sheet.getCell(lastSelectedRow, colIdx.assigned);
                    assignedCell.load("format/fill/color");
                    await context.sync();

                    if (assignedCell.format && assignedCell.format.fill) {
                        const cellColor = assignedCell.format.fill.color;
                        if (cellColor && cellColor !== '#ffffff' && cellColor !== '#000000') {
                            assignedColorMap[assignedTo] = cellColor;
                            assignedToBadge.style.backgroundColor = cellColor;
                            const r = parseInt(cellColor.substr(1, 2), 16), g = parseInt(cellColor.substr(3, 2), 16), b = parseInt(cellColor.substr(5, 2), 16);
                            assignedToBadge.style.color = ((r * 299 + g * 587 + b * 114) / 1000) > 125 ? 'black' : 'white';
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
            if (submitButton) submitButton.disabled = !currentStudentId;

            const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
            if (!panelHistory.classList.contains("hidden")) {
                if (currentStudentId) {
                    displayStudentHistory(currentStudentId);
                } else {
                    document.getElementById(CONSTANTS.HISTORY_CONTENT).innerHTML = '<p class="text-gray-500">Could not find Student ID in the selected row.</p>';
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
    if (!text || text === "N/A") return;

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
        CONSTANTS.COPY_STUDENT_ID, CONSTANTS.COPY_PRIMARY_PHONE, CONSTANTS.COPY_OTHER_PHONE,
        CONSTANTS.COPY_STUDENT_EMAIL, CONSTANTS.COPY_PERSONAL_EMAIL
    ];

    contactInfoIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('click', () => {
                const displayEl = el.querySelector('.font-semibold');
                if (displayEl) copyToClipboard(displayEl.textContent, el);
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
            if (currentGradebookLink) window.open(currentGradebookLink, '_blank');
        });
    }
}

/**
 * Finds and highlights date-related text within a comment.
 * @param {string} text The comment text.
 * @returns {string} The comment text with the date part bolded.
 */
function highlightDateInText(text) {
    const dateRegex = /\b(tomorrow|sunday|monday|tuesday|wednesday|thursday|friday|saturday|january|february|march|april|may|june|july|august|september|october|november|december)\b(\s+\d{1,2}(st|nd|rd|th)?)?/i;
    return text.replace(dateRegex, (matchedText) => `<strong class="font-bold text-gray-900">${matchedText}</strong>`);
}


/**
 * Fetches and displays the comment history for a given student ID from the "Student History" sheet.
 * @param {string} studentId The ID of the student.
 * @param {string} [searchTerm=''] An optional term to filter the comments.
 */
async function displayStudentHistory(studentId, searchTerm = '') {
    const historyContent = document.getElementById(CONSTANTS.HISTORY_CONTENT);
    historyContent.innerHTML = '<p class="text-gray-500">Loading history...</p>';
    
    const needsFetch = allComments.length === 0 || (allComments.length > 0 && String(allComments[0].studentId) !== String(studentId));

    try {
        await Excel.run(async (context) => {
            if (needsFetch) {
                const historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
                const historyRange = historySheet.getUsedRange();
                historyRange.load("values, rowIndex");
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
                    const currentId = row[idColIdx];
                    if (currentId && String(currentId) === String(studentId)) {
                        const commentText = row[commentColIdx];
                        if (commentText && String(commentText).trim() !== "") {
                            const comment = {
                                studentId: currentId,
                                rowIndex: historyRange.rowIndex + i,
                                text: commentText,
                                tag: tagColIdx !== -1 ? row[tagColIdx] : null,
                                timestamp: timestampColIdx !== -1 ? row[timestampColIdx] : null,
                                createdBy: (createdByColIdx !== -1 && row[createdByColIdx]) ? row[createdByColIdx] : 'Unknown',
                                ldaDate: null
                            };

                            if (comment.tag) {
                                const tags = String(comment.tag).split(',').map(t => t.trim());
                                const ldaTag = tags.find(t => t.toLowerCase().startsWith('lda '));
                                if (ldaTag) {
                                    const dateStr = ldaTag.substring(4);
                                    const ldaDate = new Date(dateStr);
                                    if (!isNaN(ldaDate.getTime())) {
                                        ldaDate.setHours(0, 0, 0, 0);
                                        comment.ldaDate = ldaDate;
                                    }
                                }
                            }
                            comments.push(comment);
                        }
                    }
                }
                allComments = comments;
            }
            
            let commentsToDisplay = allComments;

            if (activeFilterTags.length > 0) {
                commentsToDisplay = commentsToDisplay.filter(comment => {
                    if (!comment.tag) return false;
                    const commentTags = String(comment.tag).split(',').map(t => t.trim().split(' ')[0].toLowerCase());
                    return activeFilterTags.some(filterTag => commentTags.includes(filterTag.toLowerCase()));
                });
            }

            if (searchTerm) {
                const lowerCaseSearchTerm = searchTerm.toLowerCase();
                commentsToDisplay = commentsToDisplay.filter(comment => 
                    String(comment.text).toLowerCase().includes(lowerCaseSearchTerm) ||
                    String(comment.tag).toLowerCase().includes(lowerCaseSearchTerm) ||
                    String(comment.createdBy).toLowerCase().includes(lowerCaseSearchTerm)
                );
            }

            if (commentsToDisplay.length > 0) {
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                
                const isNoteOrDnc = (c) => c.tag && (c.tag.toLowerCase().includes('note') || c.tag.toLowerCase().includes('dnc'));

                const noteAndDncComments = commentsToDisplay.filter(isNoteOrDnc);
                const priorityComments = commentsToDisplay.filter(c => c.ldaDate && c.ldaDate >= today && !isNoteOrDnc(c));
                const regularComments = commentsToDisplay.filter(c => !isNoteOrDnc(c) && !priorityComments.includes(c));

                noteAndDncComments.sort((a, b) => (parseDate(b.timestamp) || 0) - (parseDate(a.timestamp) || 0));
                priorityComments.sort((a, b) => a.ldaDate - b.ldaDate);
                regularComments.sort((a, b) => (parseDate(b.timestamp) || 0) - (parseDate(a.timestamp) || 0));

                const sortedComments = [...noteAndDncComments, ...priorityComments, ...regularComments];
                
                const searchRegex = searchTerm ? new RegExp(`(${searchTerm.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')})`, 'gi') : null;

                let html = '<ul class="space-y-4">';
                sortedComments.forEach(comment => {
                    const isNoteOrDnc = comment.tag && (comment.tag.toLowerCase().includes('note') || comment.tag.toLowerCase().includes('dnc'));
                    const isPriority = !isNoteOrDnc && comment.ldaDate && comment.ldaDate >= today;
                    const isContacted = comment.tag && comment.tag.toLowerCase().includes('contacted');
                    
                    let bgColor = 'bg-gray-100';
                    if (isNoteOrDnc) bgColor = 'bg-gray-200';
                    else if (isPriority) bgColor = 'bg-orange-100';
                    else if (isContacted) bgColor = 'bg-yellow-100';

                    let displayText = isPriority ? highlightDateInText(comment.text) : comment.text;
                    if (searchRegex) {
                        displayText = String(displayText).replace(searchRegex, `<mark class="bg-yellow-300 p-0 m-0 rounded">$1</mark>`);
                    }

                    html += `<li class="p-3 ${bgColor} rounded-lg shadow-sm relative" data-row-index="${comment.rowIndex}">
                                <p class="text-sm text-gray-800">${displayText}</p>
                                <div class="text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center">
                                    <div class="flex items-center gap-2">`;
                    
                    if (comment.tag) {
                        String(comment.tag).split(',').map(t => t.trim()).forEach(tagName => {
                            if (tagName && tagName.toLowerCase() !== 'comment') {
                                const tagPrefix = tagName.split(' ')[0];
                                const tagInfo = availableTags.find(t => t.name.toLowerCase() === tagPrefix.toLowerCase()) || { bg: 'bg-blue-100', text: 'text-blue-800' };
                                let displayTagName = tagName;
                                if (searchRegex) displayTagName = displayTagName.replace(searchRegex, `<mark class="bg-yellow-300 p-0 m-0 rounded">$1</mark>`);
                                html += `<span class="px-2 py-0.5 font-semibold rounded-full ${tagInfo.bg} ${tagInfo.text}">${displayTagName}</span>`;
                            }
                        });
                    }
                    
                    let createdByText = comment.createdBy;
                    if (searchRegex) createdByText = createdByText.replace(searchRegex, `<mark class="bg-yellow-300 p-0 m-0 rounded">$1</mark>`);
                    html += `<span class="font-medium">${createdByText}</span></div>`;

                    let dateText = 'Unknown Time';
                    if (comment.timestamp && !isNaN(comment.timestamp) && comment.timestamp > 25569) {
                       const date = new Date((comment.timestamp - 25569) * 86400 * 1000);
                       dateText = date.toLocaleString('en-US', { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit' });
                    }
                    html += `<span>${dateText}</span></div></li>`;
                });
                html += '</ul>';
                historyContent.innerHTML = html;

                historyContent.querySelectorAll('li').forEach(li => li.addEventListener('contextmenu', showContextMenu));

            } else {
                historyContent.innerHTML = (searchTerm || activeFilterTags.length > 0)
                    ? '<p class="text-gray-500">No matching comments found.</p>'
                    : '<p class="text-gray-500">No history found for this student.</p>';
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
    const container = document.getElementById(CONSTANTS.USER_BUTTONS_CONTAINER);
    
    container.innerHTML = '';

    const userList = (settings.userProfile && settings.userProfile.userList) || [currentUserName];
    userList.forEach((user, index) => {
        const button = document.createElement('button');
        button.className = 'block w-full text-center px-4 py-2 text-sm text-gray-700 bg-gray-100 hover:bg-blue-500 hover:text-white rounded-md transition-colors';
        button.textContent = user;
        button.style.setProperty('--i', index);
        button.onclick = () => handleUserSelection(user);
        container.appendChild(button);
    });

    modal.classList.remove('hidden');
    setTimeout(() => modal.classList.add('show'), 10);
}

/**
 * Handles the user's selection from the modal, caches it, and proceeds with submission.
 * @param {string} selectedUser The name of the user that was clicked.
 */
async function handleUserSelection(selectedUser) {
    sessionCommentUser = selectedUser;
    document.getElementById(CONSTANTS.USER_SELECTION_MODAL).classList.add('hidden');
    document.getElementById(CONSTANTS.MAIN_CONTENT).classList.remove('hidden');
    document.getElementById(CONSTANTS.MAIN_CONTENT).classList.add('flex');
    
    if (!document.getElementById(CONSTANTS.TAB_DETAILS)._isInitialized) {
        initializeAddIn();
        document.getElementById(CONSTANTS.TAB_DETAILS)._isInitialized = true;
    }

    if (pendingOutreachAction) {
        const { studentId, studentName, commentText, rowIndex } = pendingOutreachAction;
        
        await addOutreachComment(studentId, studentName, commentText, sessionCommentUser);

        const lowerCommentText = commentText.toLowerCase();
        if (CONSTANTS.OUTREACH_HIGHLIGHT_TRIGGERS.some(phrase => lowerCommentText.includes(phrase))) {
            console.log(`Highlight trigger phrase found for ${studentName}. Highlighting row ${rowIndex + 1}.`);
            await applyContactedHighlight(rowIndex);
        }

        pendingOutreachAction = null;
    } else {
        const commentInput = document.getElementById(CONSTANTS.NEW_COMMENT_INPUT);
        if (commentInput && commentInput.value.trim() !== "") {
            await executeSubmitComment(sessionCommentUser);
        }
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
            
            const tagsToSave = newCommentTags.length > 0 ? newCommentTags.join(', ') : "Comment";
            if (tagCol !== -1) newRowData[tagCol] = tagsToSave;

            if (timestampCol !== -1) newRowData[timestampCol] = (new Date().getTime() / 86400000) + 25569;
            newRowData[commentCol] = commentText;

            const newRowRange = historySheet.getRangeByIndexes(newRowIndex, 0, 1, headers.length);
            newRowRange.values = [newRowData];
            
            historySheet.getUsedRange().getEntireColumn().format.autofitColumns();
            
            await context.sync();
        });

        commentInput.value = "";
        statusDisplay.textContent = "Comment added successfully!";
        
        newCommentTags = [];
        smartLdaTag = null; 
        renderTagPills('new');
        populateTagDropdown('new');
        allComments = []; // Invalidate cache
        await displayStudentHistory(currentStudentId);

        setTimeout(() => { statusDisplay.textContent = ""; }, 3000);

    } catch (error) {
        statusDisplay.textContent = error.code === 'ItemNotFound' ? `Error: "${CONSTANTS.HISTORY_SHEET}" sheet not found.` : `Error: ${error.message}`;
        console.error("Error in submitNewComment: " + error);
    }
}

// --- START: Tagging UI Functions ---

function setupTaggingUI(mode) {
    let addTagButton, tagDropdown;

    switch (mode) {
        case 'new':
            addTagButton = document.getElementById(CONSTANTS.ADD_TAG_BUTTON);
            tagDropdown = document.getElementById(CONSTANTS.TAG_DROPDOWN);
            break;
        case 'edit':
            addTagButton = document.getElementById(CONSTANTS.EDIT_ADD_TAG_BUTTON);
            tagDropdown = document.getElementById(CONSTANTS.EDIT_TAG_DROPDOWN);
            break;
        case 'filter':
            addTagButton = document.getElementById(CONSTANTS.FILTER_ADD_TAG_BUTTON);
            tagDropdown = document.getElementById(CONSTANTS.FILTER_TAG_DROPDOWN);
            break;
        default:
            return;
    }

    addTagButton.addEventListener('click', (event) => {
        event.stopPropagation();
        activeTaggingMode = mode;
        populateTagDropdown(mode);
        tagDropdown.classList.toggle('show');
    });

    document.addEventListener('click', (event) => {
        if (!tagDropdown.contains(event.target) && !addTagButton.contains(event.target)) {
            tagDropdown.classList.remove('show');
        }
    });

    if (mode !== 'filter') {
        renderTagPills(mode);
    }
    populateTagDropdown(mode);
}

function renderTagPills(mode) {
    const isEdit = mode === 'edit';
    const container = document.getElementById(isEdit ? CONSTANTS.EDIT_TAG_PILLS_CONTAINER : CONSTANTS.TAG_PILLS_CONTAINER);
    const tags = isEdit ? editCommentTags : newCommentTags;

    if (!container) return;

    container.querySelectorAll('.tag-pill').forEach(pill => pill.remove());

    const buttonContainer = container.querySelector('.relative');

    tags.forEach(tagName => {
        const tagPrefix = tagName.split(' ')[0];
        const tagInfo = availableTags.find(t => t.name === tagPrefix) || { bg: 'bg-gray-200', text: 'text-gray-800' };
        const pill = document.createElement('span');
        pill.className = `tag-pill px-2 py-1 text-xs font-semibold rounded-full flex items-center gap-1 ${tagInfo.bg} ${tagInfo.text}`;
        pill.textContent = tagName;

        const removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.innerHTML = '&times;';
        removeButton.className = 'font-bold';
        removeButton.onclick = () => removeTag(tagName, mode);

        pill.appendChild(removeButton);
        container.insertBefore(pill, buttonContainer);
    });
}


function populateTagDropdown(mode) {
    let dropdown, tags;
    switch (mode) {
        case 'new':
            dropdown = document.getElementById(CONSTANTS.TAG_DROPDOWN);
            tags = newCommentTags;
            break;
        case 'edit':
            dropdown = document.getElementById(CONSTANTS.EDIT_TAG_DROPDOWN);
            tags = editCommentTags;
            break;
        case 'filter':
            dropdown = document.getElementById(CONSTANTS.FILTER_TAG_DROPDOWN);
            tags = activeFilterTags;
            break;
        default:
            return;
    }
    
    if (!dropdown) return;
    dropdown.innerHTML = '';

    const tagsToShow = availableTags.filter(tag => !tag.hidden);

    if (tagsToShow.length === 0) {
        const noTagsItem = document.createElement('span');
        noTagsItem.className = 'block px-4 py-2 text-sm text-gray-500';
        noTagsItem.textContent = 'No more tags';
        dropdown.appendChild(noTagsItem);
        return;
    }

    tagsToShow.forEach((tag, index) => {
        const item = document.createElement('a');
        item.href = '#';
        item.className = 'block px-3 py-2 text-sm text-gray-700 hover:bg-white/20 rounded-md transition-colors';
        item.style.setProperty('--i', index);
        
        if (tag.description) {
            item.title = tag.description;
        }
        
        const pill = document.createElement('span');
        pill.className = `px-2 py-0.5 text-xs font-semibold rounded-full ${tag.bg} ${tag.text}`;
        pill.textContent = tag.name;
        item.appendChild(pill);

        item.onclick = (e) => {
            e.preventDefault();
            if (mode === 'filter') {
                toggleFilterTag(tag.name);
                dropdown.classList.remove('show');
            } else {
                if (tag.requiresDate) promptForLdaDate(mode);
                else if (tag.requiresPopup) promptForDncType(mode);
                else addTag(tag.name, mode);
            }
        };
        dropdown.appendChild(item);
    });
}

function addTag(tagName, mode) {
    const tags = mode === 'edit' ? editCommentTags : newCommentTags;
    if (!tags.some(t => t.startsWith(tagName.split(' ')[0]))) {
        tags.push(tagName);
        renderTagPills(mode);
        populateTagDropdown(mode);
        document.getElementById(mode === 'edit' ? CONSTANTS.EDIT_TAG_DROPDOWN : CONSTANTS.TAG_DROPDOWN).classList.remove('show');
    }
}

function removeTag(tagName, mode) {
    if (mode === 'edit') {
        editCommentTags = editCommentTags.filter(t => t !== tagName);
    } else {
        newCommentTags = newCommentTags.filter(t => t !== tagName);
        if (tagName === smartLdaTag) smartLdaTag = null;
    }
    renderTagPills(mode);
    populateTagDropdown(mode);
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
            addTag(`LDA ${formattedDate}`, activeTaggingMode);
            document.getElementById(CONSTANTS.DATE_PICKER_MODAL).classList.add('hidden');
        }
    });

    window.promptForLdaDate = (mode) => {
        activeTaggingMode = mode;
        selectedDate = null;
        document.getElementById('confirm-date').disabled = true;
        renderCalendar(new Date());
        document.getElementById(CONSTANTS.DATE_PICKER_MODAL).classList.remove('hidden');
        document.getElementById(mode === 'edit' ? CONSTANTS.EDIT_TAG_DROPDOWN : CONSTANTS.TAG_DROPDOWN).classList.remove('show');
    };
}


function renderCalendar(date) {
    const grid = document.getElementById('calendar-grid');
    const header = document.getElementById('month-year');
    
    grid.innerHTML = '';
    const month = date.getMonth();
    const year = date.getFullYear();

    header.textContent = `${date.toLocaleString('default', { month: 'long' })} ${year}`;

    ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach(day => {
        const dayEl = document.createElement('div');
        dayEl.className = 'font-bold text-gray-500';
        dayEl.textContent = day;
        grid.appendChild(dayEl);
    });

    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    for (let i = 0; i < firstDay; i++) grid.appendChild(document.createElement('div'));

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
            renderCalendar(date);
        };
        grid.appendChild(dayEl);
    }
}


// --- END: Date Picker Functions ---

// --- START: Smart Tag Functions ---

function handleCommentInputChange(event) {
    const text = event.target.value.toLowerCase();
    const triggerWords = ['engage', 'submit'];

    const hasTrigger = triggerWords.some(word => text.includes(word));
    const date = hasTrigger ? parseDateFromText(text) : null;

    if (date) {
        const formattedDate = `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
        const newLdaTag = `LDA ${formattedDate}`;

        if (smartLdaTag && smartLdaTag !== newLdaTag) removeTag(smartLdaTag, 'new');

        if (!smartLdaTag && !newCommentTags.some(tag => tag.startsWith('LDA'))) {
             addTag(newLdaTag, 'new');
             smartLdaTag = newLdaTag;
        }
    } else if (smartLdaTag) {
        removeTag(smartLdaTag, 'new');
    }
}

function parseDateFromText(text) {
    const now = new Date();
    text = text.toLowerCase();

    if (text.includes('tomorrow')) {
        const tomorrow = new Date();
        tomorrow.setDate(now.getDate() + 1);
        return tomorrow;
    }

    const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    for (let i = 0; i < days.length; i++) {
        if (text.includes(days[i])) {
            const targetDay = i;
            const today = now.getDay();
            let daysToAdd = targetDay - today;
            if (daysToAdd <= 0) daysToAdd += 7;
            const nextDate = new Date();
            nextDate.setDate(now.getDate() + daysToAdd);
            return nextDate;
        }
    }
    
    const monthDayRegex = /(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\s+(\d{1,2})/;
    const monthDayMatch = text.match(monthDayRegex);
    if (monthDayMatch) {
        const months = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
                         january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11 };
        const month = months[monthDayMatch[1]];
        const day = parseInt(monthDayMatch[2], 10);
        let year = now.getFullYear();
        const date = new Date(year, month, day);
        if (date < now) date.setFullYear(year + 1);
        return date;
    }

    return null;
}


// --- END: Smart Tag Functions ---

// --- START: Comment Editing and Deleting Functions ---

function setupCommentEditing() {
    const contextMenu = document.getElementById(CONSTANTS.COMMENT_CONTEXT_MENU);
    const editButton = document.getElementById(CONSTANTS.EDIT_COMMENT_BTN);
    const deleteButton = document.getElementById(CONSTANTS.DELETE_COMMENT_BTN);
    const copyButton = document.getElementById(CONSTANTS.COPY_COMMENT_BTN);
    const updateButton = document.getElementById(CONSTANTS.UPDATE_COMMENT_BUTTON);
    const cancelButton = document.getElementById(CONSTANTS.CANCEL_EDIT_BUTTON);

    document.addEventListener('click', () => contextMenu.classList.add('hidden'));

    editButton.addEventListener('click', () => {
        const commentToEdit = allComments.find(c => c.rowIndex === activeCommentRowIndex);
        if (commentToEdit) showEditCommentModal(commentToEdit);
        contextMenu.classList.add('hidden');
    });
    
    deleteButton.addEventListener('click', () => {
        document.getElementById(CONSTANTS.DELETE_CONFIRM_MODAL).classList.remove('hidden');
        contextMenu.classList.add('hidden');
    });
    
    copyButton.addEventListener('click', handleCopyComment);
    
    updateButton.addEventListener('click', handleUpdateComment);
    cancelButton.addEventListener('click', () => {
        document.getElementById(CONSTANTS.EDIT_COMMENT_MODAL).classList.add('hidden');
    });
}

function handleCopyComment() {
    const commentToCopy = allComments.find(c => c.rowIndex === activeCommentRowIndex);
    if (commentToCopy) {
        const textToCopy = commentToCopy.text;
        
        const textArea = document.createElement("textarea");
        textArea.value = textToCopy;
        textArea.style.position = "absolute";
        textArea.style.left = "-9999px";
        document.body.appendChild(textArea);
        textArea.select();
        try {
            document.execCommand('copy');
            const copyButton = document.getElementById(CONSTANTS.COPY_COMMENT_BTN);
            const originalText = copyButton.textContent;
            copyButton.textContent = "Copied!";
            setTimeout(() => {
                copyButton.textContent = originalText;
            }, 2000);
        } catch (err) {
            console.error('Failed to copy text: ', err);
        }
        document.body.removeChild(textArea);
    }
    document.getElementById(CONSTANTS.COMMENT_CONTEXT_MENU).classList.add('hidden');
}


function setupDeleteConfirmation() {
    const modal = document.getElementById(CONSTANTS.DELETE_CONFIRM_MODAL);
    const confirmBtn = document.getElementById(CONSTANTS.CONFIRM_DELETE_BUTTON);
    const cancelBtn = document.getElementById(CONSTANTS.CANCEL_DELETE_BUTTON);

    confirmBtn.addEventListener('click', handleDeleteComment);
    cancelBtn.addEventListener('click', () => modal.classList.add('hidden'));
}

function showContextMenu(event) {
    event.preventDefault();
    const contextMenu = document.getElementById(CONSTANTS.COMMENT_CONTEXT_MENU);
    const deleteButton = document.getElementById(CONSTANTS.DELETE_COMMENT_BTN);
    const commentLi = event.currentTarget;
    
    activeCommentRowIndex = parseInt(commentLi.dataset.rowIndex, 10);
    const comment = allComments.find(c => c.rowIndex === activeCommentRowIndex);

    if (comment) {
        deleteButton.disabled = comment.createdBy !== sessionCommentUser;
    } else {
        deleteButton.disabled = true;
    }

    contextMenu.style.left = `${event.clientX}px`;
    contextMenu.style.top = `${event.clientY}px`;
    contextMenu.classList.remove('hidden');
}

function showEditCommentModal(comment) {
    const modal = document.getElementById(CONSTANTS.EDIT_COMMENT_MODAL);
    const input = document.getElementById(CONSTANTS.EDIT_COMMENT_INPUT);
    const status = document.getElementById(CONSTANTS.EDIT_COMMENT_STATUS);

    input.value = comment.text;
    editCommentTags = comment.tag ? String(comment.tag).split(',').map(t => t.trim()) : [];
    status.textContent = '';
    
    renderTagPills('edit');
    populateTagDropdown('edit');
    
    modal.classList.remove('hidden');
    input.focus();
}

async function handleUpdateComment() {
    const modal = document.getElementById(CONSTANTS.EDIT_COMMENT_MODAL);
    const input = document.getElementById(CONSTANTS.EDIT_COMMENT_INPUT);
    const status = document.getElementById(CONSTANTS.EDIT_COMMENT_STATUS);
    const newText = input.value.trim();
    const newTags = editCommentTags.join(', ');

    if (!newText) {
        status.textContent = "Comment cannot be empty.";
        return;
    }
    
    status.textContent = "Updating...";

    try {
        await Excel.run(async (context) => {
            const historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
            const headersRange = historySheet.getRange("1:1").getUsedRange();
            headersRange.load("values");
            await context.sync();

            const headers = headersRange.values[0].map(h => String(h || '').toLowerCase());
            const commentCol = findColumnIndex(headers, ["comment"]);
            const tagCol = findColumnIndex(headers, ["tag"]);

            if (commentCol === -1 || tagCol === -1) throw new Error("Could not find 'Comment' or 'Tag' column.");

            const rowToUpdate = historySheet.getRangeByIndexes(activeCommentRowIndex, 0, 1, headers.length);
            rowToUpdate.getCell(0, commentCol).values = [[newText]];
            rowToUpdate.getCell(0, tagCol).values = [[newTags]];
            
            await context.sync();
        });
        
        status.textContent = "Updated successfully!";
        setTimeout(() => {
            modal.classList.add('hidden');
            allComments = []; // Invalidate cache
            displayStudentHistory(currentStudentId);
        }, 1000);

    } catch (error) {
        status.textContent = `Error: ${error.message}`;
        console.error("Error updating comment:", error);
    }
}

async function handleDeleteComment() {
    const modal = document.getElementById(CONSTANTS.DELETE_CONFIRM_MODAL);
    modal.classList.add('hidden');

    if (activeCommentRowIndex === null) return;

    try {
        await Excel.run(async (context) => {
            const historySheet = context.workbook.worksheets.getItem(CONSTANTS.HISTORY_SHEET);
            const rowToDelete = historySheet.getRangeByIndexes(activeCommentRowIndex, 0, 1, 1).getEntireRow();
            rowToDelete.delete(Excel.DeleteShiftDirection.up);
            await context.sync();
        });
        
        allComments = []; // Invalidate cache
        await displayStudentHistory(currentStudentId);

    } catch (error) {
        console.error("Error deleting comment:", error);
    } finally {
        activeCommentRowIndex = null;
    }
}


// --- END: Comment Editing and Deleting Functions ---

// --- START: DNC Functions ---

function setupDncModal() {
    const cancelButton = document.getElementById(CONSTANTS.CANCEL_DNC_BUTTON);
    cancelButton.addEventListener('click', () => {
        document.getElementById(CONSTANTS.DNC_TYPE_MODAL).classList.add('hidden');
    });
}

function promptForDncType(mode) {
    activeTaggingMode = mode;
    const modal = document.getElementById(CONSTANTS.DNC_TYPE_MODAL);
    const container = document.getElementById(CONSTANTS.DNC_OPTIONS_CONTAINER);
    container.innerHTML = '';

    const contactMethods = [
        { name: 'Phone', value: document.getElementById(CONSTANTS.PRIMARY_PHONE_DISPLAY).textContent },
        { name: 'Other Phone', value: document.getElementById(CONSTANTS.OTHER_PHONE_DISPLAY).textContent },
        { name: 'Email', value: document.getElementById(CONSTANTS.STUDENT_EMAIL_DISPLAY).textContent },
        { name: 'All', value: 'All Contact Methods' }
    ];

    contactMethods.forEach(method => {
        if (method.value && method.value !== 'N/A') {
            const button = document.createElement('button');
            button.className = 'block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border';
            button.innerHTML = `<span class="font-bold">${method.name}:</span> ${method.value}`;
            button.onclick = () => addDncTag(method.name);
            container.appendChild(button);
        }
    });

    modal.classList.remove('hidden');
    document.getElementById(mode === 'edit' ? CONSTANTS.EDIT_TAG_DROPDOWN : CONSTANTS.TAG_DROPDOWN).classList.remove('show');
}

function addDncTag(dncType) {
    let tagName = dncType === 'All' ? 'DNC' : `DNC - ${dncType}`;
    addTag(tagName, activeTaggingMode);
    document.getElementById(CONSTANTS.DNC_TYPE_MODAL).classList.add('hidden');
}


// --- END: DNC Functions ---

// --- START: Search and Filter Functions ---

function setupSearch() {
    const searchButton = document.getElementById(CONSTANTS.SEARCH_HISTORY_BUTTON);
    const searchContainer = document.getElementById(CONSTANTS.SEARCH_CONTAINER);
    const searchInput = document.getElementById(CONSTANTS.SEARCH_INPUT);
    const clearSearchButton = document.getElementById(CONSTANTS.CLEAR_SEARCH_BUTTON);

    searchButton.addEventListener('click', () => {
        searchContainer.classList.toggle('hidden');
        if (!searchContainer.classList.contains('hidden')) {
            searchInput.focus();
        }
    });

    searchInput.addEventListener('input', () => {
        if (currentStudentId) {
            displayStudentHistory(currentStudentId, searchInput.value);
        }
    });

    clearSearchButton.addEventListener('click', () => {
        searchInput.value = '';
        if (currentStudentId) {
            displayStudentHistory(currentStudentId, '');
        }
    });
}

function setupTagFilters() {
    setupTaggingUI('filter');
}

function toggleFilterTag(tagName) {
    const index = activeFilterTags.indexOf(tagName);
    if (index > -1) {
        activeFilterTags.splice(index, 1);
    } else {
        activeFilterTags.push(tagName);
    }
    renderTagFilters();
    if (currentStudentId) {
        const searchTerm = document.getElementById(CONSTANTS.SEARCH_INPUT).value;
        displayStudentHistory(currentStudentId, searchTerm);
    }
}


function renderTagFilters() {
    const container = document.getElementById(CONSTANTS.TAG_FILTER_CONTAINER);
    container.querySelectorAll('.filter-tag-button').forEach(btn => btn.remove());

    const insertTagButtonContainer = container.querySelector('.relative');

    activeFilterTags.forEach(tagName => {
        const tagInfo = availableTags.find(t => t.name === tagName) || { bg: 'bg-gray-200', text: 'text-gray-800' };
        const button = document.createElement('button');
        button.className = `filter-tag-button px-2 py-1 text-xs font-semibold rounded-full transition-all duration-150 ${tagInfo.bg} ${tagInfo.text}`;
        button.textContent = tagName;
        button.dataset.tagName = tagName;

        button.addEventListener('click', () => toggleFilterTag(tagName));
        
        container.insertBefore(button, insertTagButtonContainer);
    });
}

// --- END: Search and Filter Functions ---


// --- START: Code moved from commands.js ---

function errorHandler(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}

function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) return dateValue;
    if (typeof dateValue === 'number' && dateValue > 25569) return new Date((dateValue - 25569) * 86400 * 1000);
    if (typeof dateValue === 'string') {
        const parsed = new Date(dateValue);
        if (!isNaN(parsed.getTime())) return parsed;
    }
    return null;
}

function jsDateToExcelDate(date) {
    return (date.getTime() / 86400000) + 25569;
}

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
            const colCount = Math.abs(studentNameColIndex - outreachColIndex) + 1;

            const highlightRange = sheet.getRangeByIndexes(rowIndex, startCol, 1, colCount);
            highlightRange.format.fill.color = "yellow";
            await context.sync();
        });
    } catch (error) {
        errorHandler(error);
    }
}

async function onWorksheetChanged(eventArgs) {
    await Excel.run(async (context) => {
        if (eventArgs.source !== Excel.EventSource.local || (eventArgs.changeType !== "CellEdited" && eventArgs.changeType !== "RangeEdited")) {
            return;
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const changedRange = sheet.getRange(eventArgs.address);
        const headerRange = sheet.getRange("1:1").getUsedRange(true);

        changedRange.load("address, rowIndex, columnIndex, rowCount, columnCount, values, valuesBefore");
        headerRange.load("values, columnCount");
        await context.sync();

        const headers = (headerRange.values[0] || []).map(h => String(h || '').toLowerCase());
        const outreachColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.outreach);

        if (outreachColIndex === -1 || changedRange.rowIndex === 0) {
            return; // No outreach column or change is in the header
        }

        // Check if the changed range intersects with the outreach column at all
        const outreachColumnAffected = changedRange.columnIndex <= outreachColIndex && (changedRange.columnIndex + changedRange.columnCount - 1) >= outreachColIndex;
        if (!outreachColumnAffected) {
            return;
        }
        
        // This offset is crucial for finding the correct value in multi-column changes
        const outreachColumnOffset = outreachColIndex - changedRange.columnIndex;

        // Load all data for the affected rows at once to be efficient
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
            return; // Can't identify student
        }

        // Now, loop through each row that was changed
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
                    if (!sessionCommentUser) {
                        // If no user is selected, save the *first* action and prompt.
                        // We won't process subsequent changes in this batch to avoid multiple prompts.
                        if (!pendingOutreachAction) { 
                            console.log("Outreach change detected, but no user is selected. Prompting.");
                            pendingOutreachAction = { studentId, studentName, commentText: newValue, rowIndex };
                            promptForUserAndSubmit();
                        }
                        return; // Stop processing further changes in this event until user is selected
                    }
                    
                    // If user is selected, process the comment immediately
                    await addOutreachComment(studentId, studentName, newValue, sessionCommentUser);
                    const lowerNewValue = newValue.toLowerCase();
                    if (CONSTANTS.OUTREACH_HIGHLIGHT_TRIGGERS.some(phrase => lowerNewValue.includes(phrase))) {
                        console.log(`Highlight trigger phrase found for ${studentName}. Highlighting row ${rowIndex + 1}.`);
                        await applyContactedHighlight(rowIndex);
                    }
                }
            }
        }
    }).catch(errorHandler);
}


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
                                break;
                            }
                        }
                    }
                }
            }

            const now = new Date();
            const excelNow = jsDateToExcelDate(now);

            if (todaysCommentRowIndex !== -1) {
                const commentCell = historySheet.getCell(todaysCommentRowIndex, commentCol);
                const timestampCell = historySheet.getCell(todaysCommentRowIndex, timestampCol);
                
                commentCell.values = [[commentText]];
                timestampCell.values = [[excelNow]];
                timestampCell.numberFormat = [["M/D/YYYY h:mm AM/PM"]];

            } else {
                const newRowData = new Array(historyHeaders.length).fill("");
                newRowData[idCol] = studentId;
                if (studentCol !== -1) newRowData[studentCol] = studentName;
                if (createdByCol !== -1) newRowData[createdByCol] = commentingUser;
                
                let tagsToSave = "Outreach";
                const date = parseDateFromText(commentText);
                if (date) {
                    const formattedDate = `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
                    tagsToSave += `, LDA ${formattedDate}`;
                }

                if (CONSTANTS.OUTREACH_HIGHLIGHT_TRIGGERS.some(phrase => commentText.toLowerCase().includes(phrase))) {
                    tagsToSave += ', Contacted';
                }

                newRowData[tagCol] = tagsToSave;
                newRowData[timestampCol] = excelNow;
                newRowData[commentCol] = commentText;
                
                const newRowIndex = historyRange.rowIndex + historyRange.rowCount;
                const newRowRange = historySheet.getRangeByIndexes(newRowIndex, 0, 1, historyHeaders.length);
                newRowRange.values = [newRowData];
                
                historySheet.getCell(newRowIndex, timestampCol).numberFormat = [["M/D/YYYY h:mm AM/PM"]];
            }
            
            historySheet.getUsedRange().format.autofitColumns();
            await context.sync();

            if (studentId && String(studentId) === String(currentStudentId)) {
                const panelHistory = document.getElementById(CONSTANTS.PANEL_HISTORY);
                if (panelHistory && !panelHistory.classList.contains("hidden")) {
                    allComments = []; // Invalidate cache
                    setTimeout(() => displayStudentHistory(currentStudentId), 100);
                }
            }

        } catch (error) {
            errorHandler(error);
        }
    });
}
