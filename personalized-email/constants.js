export const DOM_IDS = {
    // Setup
    SETUP_WIZARD: 'setup-wizard',
    POWER_AUTOMATE_URL: 'power-automate-url',
    CREATE_CONNECTION_BUTTON: 'create-connection-button',
    SETUP_STATUS: 'setup-status',
    
    // Composer
    EMAIL_COMPOSER: 'email-composer',
    RECIPIENT_LIST: 'recipient-list',
    CUSTOM_SHEET_CONTAINER: 'custom-sheet-container',
    CUSTOM_SHEET_NAME: 'custom-sheet-name',
    EDITOR_CONTAINER: 'editor-container',
    STANDARD_PARAMETER_BUTTONS: 'standard-parameter-buttons',
    CUSTOM_PARAMETERS_SECTION: 'custom-parameters-section',
    CUSTOM_PARAMETER_BUTTONS: 'custom-parameter-buttons',
    RANDOMIZE_PARAMETER_BUTTON: 'randomize-parameter-button',
    CONDITION_PARAMETER_BUTTON: 'condition-parameter-button',
    CREATE_CUSTOM_PARAM_BUTTON: 'create-custom-param-button',
    SEND_EMAIL_BUTTON: 'send-email-button',
    SHOW_EXAMPLE_BUTTON: 'show-example-button',
    SHOW_PAYLOAD_BUTTON: 'show-payload-button',
    TEMPLATES_BUTTON: 'templates-button',
    STATUS: 'status',

    // Pillbox Inputs
    FROM_CONTAINER: 'email-from-container',
    FROM_INPUT: 'email-from-input',
    SUBJECT_CONTAINER: 'email-subject-container',
    SUBJECT_INPUT: 'email-subject-input',
    CC_CONTAINER: 'email-cc-container',
    CC_INPUT: 'email-cc-input',

    // Modals
    EXAMPLE_MODAL: 'example-modal',
    EXAMPLE_FROM: 'example-from',
    EXAMPLE_TO: 'example-to',
    EXAMPLE_CC: 'example-cc',
    EXAMPLE_SUBJECT: 'example-subject',
    EXAMPLE_BODY: 'example-body',
    CLOSE_EXAMPLE_MODAL_BUTTON: 'close-example-modal-button',

    PAYLOAD_MODAL: 'payload-modal',
    PAYLOAD_CONTENT: 'payload-content',
    CLOSE_PAYLOAD_MODAL_BUTTON: 'close-payload-modal-button',
    
    SEND_CONFIRM_MODAL: 'send-confirm-modal',
    SEND_CONFIRM_MESSAGE: 'send-confirm-message',
    CANCEL_SEND_BUTTON: 'cancel-send-button',
    CONFIRM_SEND_BUTTON: 'confirm-send-button',

    TEMPLATES_MODAL: 'templates-modal',
    TEMPLATES_LIST_CONTAINER: 'templates-list-container',
    SAVE_CURRENT_TEMPLATE_BUTTON: 'save-current-template-button',
    CLOSE_TEMPLATES_MODAL_BUTTON: 'close-templates-modal-button',

    SAVE_TEMPLATE_MODAL: 'save-template-modal',
    TEMPLATE_NAME: 'template-name',
    TEMPLATE_AUTHOR: 'template-author',
    SAVE_TEMPLATE_STATUS: 'save-template-status',
    CANCEL_SAVE_TEMPLATE_BUTTON: 'cancel-save-template-button',
    CONFIRM_SAVE_TEMPLATE_BUTTON: 'confirm-save-template-button',

    CUSTOM_PARAM_MODAL: 'custom-param-modal',
    CUSTOM_PARAM_MODAL_TITLE: 'custom-param-modal-title',
    PARAM_EDIT_ID: 'param-edit-id',
    PARAM_NAME: 'param-name',
    PARAM_SOURCE_COLUMN: 'param-source-column',
    PARAM_DEFAULT_VALUE: 'param-default-value',
    PARAM_MAPPING_CONTAINER: 'param-mapping-container',
    ADD_MAPPING_BUTTON: 'add-mapping-button',
    SAVE_PARAM_STATUS: 'save-param-status',
    CANCEL_CUSTOM_PARAM_BUTTON: 'cancel-custom-param-button',
    SAVE_CUSTOM_PARAM_BUTTON: 'save-custom-param-button',
    MANAGE_CUSTOM_PARAMS_BUTTON: 'manage-custom-params-button',
    
    MANAGE_CUSTOM_PARAMS_MODAL: 'manage-custom-params-modal',
    MANAGE_PARAMS_LIST: 'manage-params-list',
    CLOSE_MANAGE_PARAMS_BUTTON: 'close-manage-params-button'
};

export const SETTINGS_KEYS = {
    CONNECTIONS: 'connections',
    EMAIL_TEMPLATES: 'emailTemplates',
    CUSTOM_PARAMS: 'customEmailParameters'
};

export const STANDARD_PARAMETERS = ['FirstName', 'LastName', 'StudentName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut'];

/**
 * A map of standard parameter names to possible column headers in the Excel sheet.
 * All headers are checked in lowercase.
 */
export const COLUMN_MAPPINGS = {
    StudentName: ["studentname", "student name", "student"],
    StudentEmail: ["student email", "school email", "email"],
    PersonalEmail: ["personal email", "otheremail"],
    Grade: ["grade", "course grade"],
    DaysOut: ["days out", "daysout"]
};

