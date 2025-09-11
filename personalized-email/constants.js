// V-1.3 - 2025-09-11 - 11:37 AM EDT

/**
 * @fileoverview This file contains constants and configurations for the Personalized Email add-in.
 * Organizing these values here makes the main script cleaner and easier to manage.
 */

// --- CORE SETTINGS KEYS ---
// These keys are used to store and retrieve data from the workbook settings.

// Key for storing email templates created by the user.
export const EMAIL_TEMPLATES_KEY = "emailTemplates";

// Key for storing custom parameters defined by the user.
export const CUSTOM_PARAMS_KEY = "customEmailParameters";


// --- PARAMETERS ---
// These are the placeholders that can be inserted into emails.

// Standard, built-in parameters that are always available.
export const standardParameters = [
    'FirstName', 
    'LastName', 
    'StudentName', 
    'StudentEmail', 
    'PersonalEmail', 
    'Grade', 
    'DaysOut', 
    'Assigned'
];

// --- QUILL EDITOR CONFIGURATION ---
// Settings for the Quill rich text editor.
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

// --- DATA MAPPING ---
// Defines how to map spreadsheet columns to data fields.

// Maps standard parameter names to possible column headers in the Excel sheet.
// The names are case-insensitive.
export const COLUMN_MAPPINGS = {
    StudentName: ["studentname", "student name"],
    StudentEmail: ["student email", "school email", "email"],
    PersonalEmail: ["personal email", "otheremail"],
    Grade: ["grade", "course grade"],
    DaysOut: ["days out", "daysout"],
    Assigned: ["assigned"]
};


// --- POWER AUTOMATE CONFIGURATION ---
// Defines the structure of the data sent to the Power Automate flow.

// The JSON schema for the payload sent to Power Automate.
// This helps ensure the data is in the correct format.
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

