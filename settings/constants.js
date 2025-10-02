// Timestamp: 2025-10-02 10:04 AM | Version: 1.0.0
const CONSTANTS = {
    // Keys for Office.context.document.settings
    SETTINGS_KEYS: {
        APP: "studentRetentionSettings"
    },
    // Names of worksheets used in the add-in
    SHEET_NAMES: {
        MASTER_LIST: "Master List"
    },
    // DOM element IDs used in settings.html
    DOM: {
        IDS: {
            // Buttons
            saveButton: "save-button",
            resetButton: "reset-button",
            addUserButton: "add-user-button",
            cancelAddUserButton: "cancel-add-user-button",
            saveNewUserButton: "save-new-user-button",

            // User Management
            userListContainer: "user-list-container",
            
            // Task Pane Settings
            smartNavigationToggle: "smart-navigation-toggle",

            // LDA Report & Import Settings
            daysOutFilter: "days-out-filter",
            includeFailingList: "include-failing-list",
            includeLdaTagFollowup: "include-lda-tag-followup",
            treatEmptyGradesAsZero: "treat-empty-grades-as-zero",
            hideLeftoverColumns: "hide-leftover-columns",

            // Columns Management
            includedColumns: "included-columns",
            availableColumns: "available-columns",
            columnsLoader: "columns-loader",
            
            // Modals & Status
            addUserModal: "add-user-modal",
            newUserFirstName: "new-user-first-name",
            newUserLastName: "new-user-last-name",
            addUserStatus: "add-user-status",
            status: "status"
        },
        CLASSES: {
            // General
            hidden: "hidden",
            
            // Status Messages
            statusMessage: "status-message",
            statusSuccess: "status-success",
            statusError: "status-error",
            visible: "visible",

            // User Item
            userItem: "user-item",
            userNameContainer: "user-name-container",
            userNameDisplay: "user-name-display",
            userNameInput: "user-name-input",
            userItemActions: "user-item-actions",
            iconButton: "icon-button",
            editButton: "edit-button",
            saveButton: "save-button",
            cancelButton: "cancel-button",
            removeButton: "remove-button",

            // Columns
            columnItem: "column-item",
            sortableGhost: "sortable-ghost"
        }
    },
    // Default settings structure
    DEFAULT_SETTINGS: {
        createlda: {
            daysOutFilter: 6,
            includeFailingList: true,
            includeLdaTagFollowup: true,
            hideLeftoverColumns: true,
            treatEmptyGradesAsZero: false,
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        },
        userProfile: {
            name: "",
            userList: []
        },
        taskpane: {
            smartNavigation: true
        }
    }
};