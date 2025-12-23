/*
 * This file contains shared constants and utility functions used across the add-in.
 */

export const CONSTANTS = {
    // NOTE: "Student ID" and "Student Number" are treated as distinct values.
    // "Student ID" refers to the Canvas ID, used for creating gradebook links.
    // "Student Number" (and "Student Identifier") refers to the internal school ID.
    STUDENT_NAME_COLS: ["studentname", "student name","student"],
    OUTREACH_COLS: ["outreach"],
    STUDENT_ID_COLS: ["student id", "systudentid", "id"],
    STUDENT_NUMBER_COLS: ["studentnumber", "student identifier", "student number"],
    MASTER_LIST_SHEET: "Master List",
    HISTORY_SHEET: "Student History",
    SETTINGS_KEY: "studentRetentionSettings", // Key for document settings
    COLUMN_MAPPINGS: {
        course: ["course"],
        courseId: ["course id"],
        courseLastAccess: ["course last access"],
        currentScore: ["current score", "grade", "course grade"],
        grade: ["grade", "course grade", "Grade", "Course Grade", "grades", "Grades"],
        gradeBook: ["grade book", "gradebook", "Grade Book", "Gradebook", "GradeBook", "gradebooklink", "GradebookLink", "Gradelink"],
        daysOut: ["days out", "Days Out"],
        lastLda: ["lda", "last lda", "LDA", "Last LDA"],
        assigned: ["assigned", "Assigned"],
        programVersion: ["programversion", "program version"],
        courseMissingAssignments: ["course missing assignments"],
        courseZeroAssignments: ["course zero assignments"],
        primaryPhone: ["primary phone", "phone", "phonenumber", "contact number", "Phone", "Phone Number", "Contact Number"],
        otherPhone: ["other phone", "otherphone", "Other Phone", "OtherPhone"],
        PersonalEmail: ["personal email", "otheremail", "Personal Email"],
        StudentEmail: ["student email", "school email", "email", "Email","Student Email"]
    }
};

/**
 * Generic error handler for Excel.run calls.
 * @param {any} error The error object.
 */
export function errorHandler(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}

/**
 * Gets the settings object from document settings, ensuring it's the latest version.
 * @returns {Promise<object>} A promise that resolves with the parsed settings object.
 */
export async function getSettings() {
    // First, refresh the settings from the document to ensure we have the latest version.
    await new Promise((resolve) => {
        Office.context.document.settings.refreshAsync(asyncResult => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error refreshing settings: " + asyncResult.error.message);
                // Even if refresh fails, we proceed with the cached version.
            } else {
                console.log("Settings refreshed successfully.");
            }
            resolve();
        });
    });

    // Now, get the settings value.
    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEY);
    const defaults = {
        createlda: {
            daysOutFilter: 6,
            includeFailingList: true,
            hideLeftoverColumns: true,
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        }
    };

    if (settingsString) {
        try {
            const settings = JSON.parse(settingsString);
            settings.createlda = { ...defaults.createlda, ...(settings.createlda || {}) };
            return settings;
        } catch (e) {
            console.error("Error parsing settings, returning defaults:", e);
            return defaults;
        }
    }
    return defaults;
}


/**
 * Parses a date value from various possible formats (Date object, string, Excel serial number).
 * @param {*} dateValue The value to parse.
 * @returns {Date|null} A valid Date object or null.
 */
export function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
        // ExcelJS can return Date objects for .xlsx
        return dateValue;
    }
    if (typeof dateValue === 'number') {
        // Excel serial date number
        // Check for a reasonable range to avoid treating random numbers as dates
        if (dateValue > 25569) { // Corresponds to 1970-01-01
            return new Date((dateValue - 25569) * 86400 * 1000);
        }
    }
    if (typeof dateValue === 'string') {
        // Try parsing common date formats
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
export function jsDateToExcelDate(date) {
    return (date.getTime() / 86400000) + 25569;
}


/**
 * Helper to normalize names from "Last, First" or "First Last" to "first last"
 * for consistent matching.
 * @param {string} name The name to normalize.
 * @returns {string} The normalized name.
 */
export const normalizeName = (name) => {
    if (!name || typeof name !== 'string') return '';
    name = name.trim().toLowerCase();
    if (name.includes(',')) {
        const parts = name.split(',').map(part => part.trim());
        if (parts.length > 1) {
            return `${parts[1]} ${parts[0]}`;
        }
    }
    return name;
};

/**
 * Helper to format names to "Last, First" format.
 * @param {string} name The name to format.
 * @returns {string} The formatted name.
 */
export const formatToLastFirst = (name) => {
    if (!name || typeof name !== 'string') return '';
    name = name.trim();
    if (name.includes(',')) {
        // Already "Last, First"
        return name.split(',').map(p => p.trim()).join(', ');
    }
    const parts = name.split(' ').filter(p => p);
    if (parts.length > 1) {
        const lastName = parts.pop();
        const firstName = parts.join(' ');
        return `${lastName}, ${firstName}`;
    }
    return name;
};


/**
 * Converts a data URL to an ArrayBuffer.
 * @param {string} dataUrl The data URL.
 * @returns {ArrayBuffer}
 */
export function dataUrlToArrayBuffer(dataUrl) {
    const base64String = dataUrl.substring(dataUrl.indexOf(',') + 1);
    const binaryString = window.atob(base64String);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

/**
 * A robust CSV row parser that handles quoted fields.
 */
export function parseCsvRow(row) {
    const cells = [];
    let inQuotes = false;
    let cell = '';
    for (let i = 0; i < row.length; i++) {
        const char = row[i];
        if (char === '"') {
            if (inQuotes && row[i+1] === '"') {
                cell += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            cells.push(cell);
            cell = '';
        } else {
            cell += char;
        }
    }
    cells.push(cell);
    return cells;
}

/**
 * Finds the index of a column by checking against a list of possible names.
 * Includes a check to ensure possibleNames is an array.
 */
export function findColumnIndex(headers, possibleNames) {
    if (!Array.isArray(possibleNames)) {
        console.error("[DEBUG] findColumnIndex received non-array for possibleNames:", possibleNames);
        return -1;
    }
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}
