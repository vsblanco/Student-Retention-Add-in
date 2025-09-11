// V-1.7 - 2025-09-11 - 12:29 PM EDT
/**
 * @fileoverview This file contains all the constant configurations used in the Personalized Email add-in.
 * It is organized from most important/frequently accessed at the top to least at the bottom.
 */

// --- Core Application Keys ---
// These keys are used for storing and retrieving settings within the Excel workbook.
export const EMAIL_TEMPLATES_KEY = "emailTemplates";
export const CUSTOM_PARAMS_KEY = "customEmailParameters";

// --- Parameter Configurations ---
// Standard, built-in parameters available for personalization.
export const standardParameters = ['FirstName', 'LastName', 'StudentName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut', 'Assigned'];

// --- Column Mapping Configurations ---
// Maps standard parameter names to possible column headers in the Excel sheet.
// This allows for flexibility in sheet formatting. The first name found in the headers array will be used.
export const COLUMN_MAPPINGS = {
    StudentName: ["studentname", "student name"],
    StudentEmail: ["student email", "school email", "email"],
    PersonalEmail: ["personal email", "otheremail"],
    Grade: ["grade", "course grade"],
    DaysOut: ["days out", "daysout"],
    Assigned: ["assigned"]
};

// --- Quill Editor Configuration ---
// Defines the settings for the Quill rich text editor, including the theme and toolbar options.
export const QUILL_EDITOR_CONFIG = {
    theme: 'snow',
    modules: {
        toolbar: [
            ['bold', 'italic', 'underline'],
            [{'list': 'ordered'}, {'list': 'bullet'}],
            [{'color': []}, {'background': []}],
            ['link']
        ]
    }
};

// --- Power Automate Payload Schema ---
// The JSON schema for the data that will be sent to the Power Automate flow.
// This helps ensure the data structure is correct.
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

// --- Custom Parameter Mapping Operators ---
// Defines the available operators for value mapping in custom parameters.
export const MAPPING_OPERATORS = [
    { value: 'eq', text: 'is equal to' },
    { value: 'neq', text: 'is not equal to' },
    { value: 'contains', text: 'contains' },
    { value: 'does_not_contain', text: 'does not contain' },
    { value: 'starts_with', text: 'starts with' },
    { value: 'ends_with', text: 'ends with' },
];

