// V-5.0 - 2025-09-30 - 5:10 PM EDT
// -- CORE CONFIGURATION --
// These are the primary keys used to store data within the Excel workbook's settings.
export const EMAIL_TEMPLATES_KEY = "emailTemplates";
export const CUSTOM_PARAMS_KEY = "customEmailParameters";

// -- PARAMETERS --
// Standard, built-in parameters that are always available.
export const standardParameters = ['FirstName', 'LastName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut'];

// -- QUILL EDITOR SETTINGS --
// Configuration for the main rich text editor.
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

// Configuration for the smaller rich text editors used in the custom parameter modal.
export const MINI_QUILL_EDITOR_CONFIG = {
    theme: 'snow',
    modules: {
        toolbar: [
            ['bold', 'italic', 'underline', 'link']
        ]
    }
};

// -- DATA MAPPING --
// Maps internal data keys to possible column header names in the Excel sheet.
export const COLUMN_MAPPINGS = {
    StudentName: ["studentname", "student name", "Student Name"],
    StudentEmail: ["student email", "school email", "email", "Student Email", "StudentEmail"],
    PersonalEmail: ["personal email", "otheremail", "Personal Email", "PersonalEmail"],
    Grade: ["grade", "course grade", "Grade", "Course Grade", "grades", "Grades"],
    DaysOut: ["days out", "daysout", "Days Out"],
    Assigned: ['assigned', 'advisor', 'counselor', 'assigned to', "Assigned", "Advisor"],
    Tags: ['history', 'tags', 'tag', "Tag", "Tags"],
    StudentIdentifier: ['student identifier','studentnumber', "Student Identifier", "Student Number"],
    Outreach: ['outreach', "Outreach"]
};

// Defines the operators available for creating custom parameter mapping rules.
export const MAPPING_OPERATORS = [
    { value: 'eq', text: 'is equal to' },
    { value: 'neq', text: 'is not equal to' },
    { value: 'contains', text: 'contains' },
    { value: 'does_not_contain', text: 'does not contain' },
    { value: 'starts_with', text: 'starts with' },
    { value: 'ends_with', text: 'ends with' },
    { value: 'gt', text: 'is greater than' },
    { value: 'lt', text: 'is less than' },
    { value: 'gte', text: 'is greater than or equal to' },
    { value: 'lte', text: 'is less than or equal to' }
];

// -- UI STYLING CONSTANTS --
// Contains Tailwind CSS class strings for different UI elements.

export const PARAMETER_BUTTON_STYLES = {
    // Style for standard, built-in parameters.
    standard: 'px-2 py-1 bg-gray-200 text-gray-700 text-xs rounded hover:bg-gray-300',
    // Style for simple custom parameters (no mappings).
    custom: 'px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200',
    // Style for custom parameters that have mapping rules.
    mapped: 'px-2 py-1 bg-purple-100 text-purple-800 text-xs rounded hover:bg-purple-200',
    // Style for custom parameters that have mappings which include other nested parameters.
    nested: 'px-2 py-1 bg-rose-100 text-rose-800 text-xs rounded hover:bg-rose-200',
    // Style for custom parameters that are powered by a Custom Script.
    script: 'px-2 py-1 bg-gray-500 text-white text-xs rounded hover:bg-gray-600'
};


// -- POWER AUTOMATE SCHEMA --
// The JSON schema for the payload that will be sent to the Power Automate flow.
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

