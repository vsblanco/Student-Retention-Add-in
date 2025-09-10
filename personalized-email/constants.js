export const DOM_IDS = {
    // Views
    SETUP_WIZARD: 'setup-wizard',
    EMAIL_COMPOSER: 'email-composer',

    // Setup Wizard
    POWER_AUTOMATE_URL: 'power-automate-url',
    CREATE_CONNECTION_BUTTON: 'create-connection-button',
    SETUP_STATUS: 'setup-status',

    // Email Composer
    RECIPIENT_LIST: 'recipient-list',
    CUSTOM_SHEET_CONTAINER: 'custom-sheet-container',
    CUSTOM_SHEET_NAME: 'custom-sheet-name',
    TEMPLATES_BUTTON: 'templates-button',
    STATUS: 'status',

    // Email Fields
    FROM_CONTAINER: 'email-from-container',
    FROM_INPUT: 'email-from-input',
    SUBJECT_CONTAINER: 'email-subject-container',
    SUBJECT_INPUT: 'email-subject-input',
    CC_CONTAINER: 'email-cc-container',
    CC_INPUT: 'email-cc-input',
    EDITOR_CONTAINER: 'editor-container',

    // Parameters
    STANDARD_PARAMETER_BUTTONS: 'standard-parameter-buttons',
    CUSTOM_PARAMETERS_SECTION: 'custom-parameters-section',
    CUSTOM_PARAMETER_BUTTONS: 'custom-parameter-buttons',
    CREATE_CUSTOM_PARAM_BUTTON: 'create-custom-param-button',
    RANDOMIZE_PARAMETER_BUTTON: 'randomize-parameter-button',
    CONDITION_PARAMETER_BUTTON: 'condition-parameter-button',

    // Action Buttons
    SEND_EMAIL_BUTTON: 'send-email-button',
    SHOW_EXAMPLE_BUTTON: 'show-example-button',
    SHOW_PAYLOAD_BUTTON: 'show-payload-button',

    // Modals
    EXAMPLE_MODAL: 'example-modal',
    EXAMPLE_FROM: 'example-from',
    EXAMPLE_TO: 'example-to',
    EXAMPLE_CC: 'example-cc',
    EXAMPLE_SUBJECT: 'example-subject',
    EXAMPLE_BODY: 'example-body',
    CLOSE_EXAMPLE_MODAL_BUTTON: 'close-example-modal-button',

    PAYLOAD_MODAL: 'payload-modal',
    PAYLOAD_MODAL_TITLE: 'payload-modal-title',
    PAYLOAD_CONTENT: 'payload-content',
    SCHEMA_CONTENT: 'schema-content',
    TOGGLE_PAYLOAD_SCHEMA_BUTTON: 'toggle-payload-schema-button',
    CLOSE_PAYLOAD_MODAL_BUTTON: 'close-payload-modal-button',

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
    CLOSE_MANAGE_PARAMS_BUTTON: 'close-manage-params-button',

    SEND_CONFIRM_MODAL: 'send-confirm-modal',
    SEND_CONFIRM_MESSAGE: 'send-confirm-message',
    CANCEL_SEND_BUTTON: 'cancel-send-button',
    CONFIRM_SEND_BUTTON: 'confirm-send-button',
};

export const SETTINGS_KEYS = {
    CONNECTIONS: "connections",
    EMAIL_TEMPLATES: "emailTemplates",
    CUSTOM_PARAMS: "customEmailParameters",
};

export const STANDARD_PARAMETERS = ['FirstName', 'LastName', 'StudentName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut'];

export const PAYLOAD_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "from": { "type": "string" },
            "to": { "type": "string" },
            "cc": { "type": "string" },
            "subject": { "type": "string" },
            "body": { "type": "string" }
        },
        "required": ["from", "to", "subject", "body"]
    }
};
