// Constants for personalized email feature
export const EMAIL_TEMPLATES_KEY = "emailTemplates";
export const CUSTOM_PARAMS_KEY = "customEmailParameters";

// Standard, built-in parameters that are always available
export const standardParameters = ['FirstName', 'LastName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut', 'Assigned'];

// Special built-in parameters (more advanced, dynamically generated)
export const specialParameters = ['MissingAssignmentsList', 'DaysLeft', 'Salutation'];

// Configuration for the main rich text editor
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

// Configuration for smaller rich text editors (value mapping result boxes)
export const MINI_QUILL_EDITOR_CONFIG = {
    theme: 'snow',
    modules: {
        toolbar: [
            ['bold', 'italic', 'underline', 'link']
        ]
    }
};

// Maps internal data keys to possible column header names in the Excel sheet
export const COLUMN_MAPPINGS = {
    StudentName: ["studentname", "student name", "Student Name", "StudentName"],
    StudentEmail: ["student email", "school email", "email", "Student Email", "StudentEmail"],
    PersonalEmail: ["personal email", "otheremail", "Personal Email", "PersonalEmail"],
    Grade: ["grade", "course grade", "Grade", "Course Grade", "grades", "Grades"],
    DaysOut: ["days out", "daysout", "Days Out"],
    Assigned: ['assigned', 'advisor', 'counselor', 'assigned to', "Assigned", "Advisor"],
    Tags: ['history', 'tags', 'tag', "Tag", "Tags"],
    StudentIdentifier: ['student identifier','studentnumber', "Student Identifier", "Student Number"],
    Outreach: ['outreach', "Outreach"],
    GradeBook: ['grade book', 'gradebook', 'Grade Book', 'GradeBook'],
    Assignment: ['assignment', 'Assignment']
};

// Defines the operators available for creating custom parameter mapping rules
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

// Tailwind CSS class strings for different UI elements
export const PARAMETER_BUTTON_STYLES = {
    standard: 'px-2 py-1 bg-gray-200 text-gray-700 text-xs rounded hover:bg-gray-300',
    special: 'px-2 py-1 bg-orange-100 text-orange-800 text-xs rounded hover:bg-orange-200',
    custom: 'px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200',
    mapped: 'px-2 py-1 bg-purple-100 text-purple-800 text-xs rounded hover:bg-purple-200',
    nested: 'px-2 py-1 bg-rose-100 text-rose-800 text-xs rounded hover:bg-rose-200'
};

// The JSON schema for the payload that will be sent to the Power Automate flow
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
