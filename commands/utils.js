/*
 * This file contains shared constants and utility functions used across the add-in.
 */

export const CONSTANTS = {
    // NOTE: "Student ID" and "Student Number" are treated as distinct values.
    // "Student ID" refers to the Canvas ID, used for creating gradebook links.
    // "Student Number" (and "Student Identifier") refers to the internal school ID.
    STUDENT_NAME_COLS: ["studentname", "student name","student"],
    OUTREACH_COLS: ["outreach"],
    STUDENT_ID_COLS: ["student id"],
    STUDENT_NUMBER_COLS: ["studentnumber", "student identifier"],
    MASTER_LIST_SHEET: "Master List",
    HISTORY_SHEET: "Student History",
    SETTINGS_KEY: "studentRetentionSettings", // Key for document settings
    COLUMN_MAPPINGS: {
        course: ["course"],
        courseId: ["course id"],
        courseLastAccess: ["course last access"],
        currentScore: ["current score", "grade", "course grade"],
        grade: ["grade", "course grade"],
        gradeBook: ["grade book", "gradebook"],
        daysOut: ["days out"],
        lastLda: ["lda", "last lda"],
        assigned: ["assigned"],
        programVersion: ["programversion", "program version"],
        courseMissingAssignments: ["course missing assignments"],
        courseZeroAssignments: ["course zero assignments"]
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
 * Parses a date value from various possible formats (Date object, string, Excel serial number),
 * correcting for timezone issues.
 * @param {*} dateValue The value to parse.
 * @returns {Date|null} A valid Date object or null.
 */
export function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
        // ExcelJS can return Date objects for .xlsx. These are usually timezone-aware.
        return dateValue;
    }
    if (typeof dateValue === 'number') {
        // Excel serial date number
        if (dateValue > 25569) { // Corresponds to 1970-01-01
            // This calculation gives milliseconds since Unix epoch (UTC)
            const utcMilliseconds = (dateValue - 25569) * 86400 * 1000;
            const date = new Date(utcMilliseconds);
            // We need to create a new date that represents the same "wall clock" time in the local timezone.
            // We do this by adding the timezone offset that was implicitly subtracted by `new Date()`.
            const correctedDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60000);
            return correctedDate;
        }
    }
    if (typeof dateValue === 'string') {
        // For strings like "MM/DD/YYYY", treat them as local time.
        // The jsDateToExcelDate function will handle converting this to the correct serial number.
        const parsed = new Date(dateValue.replace(/-/g, '/'));
        if (!isNaN(parsed.getTime())) {
            return parsed;
        }
    }
    return null;
}


/**
 * Converts a JavaScript Date object to an Excel serial date number, preserving the "wall clock" time.
 * @param {Date} date The JavaScript Date object.
 * @returns {number} The Excel serial date number.
 */
export function jsDateToExcelDate(date) {
    // Create a UTC timestamp using the local components of the date (year, month, day, time).
    // This effectively strips the timezone information and preserves the "wall clock" time.
    const utcTimestamp = Date.UTC(
        date.getFullYear(),
        date.getMonth(),
        date.getDate(),
        date.getHours(),
        date.getMinutes(),
        date.getSeconds(),
        date.getMilliseconds()
    );
    // Convert UTC timestamp to Excel serial number.
    // The 25569 is the day difference between Excel epoch (1900) and Unix epoch (1970).
    return (utcTimestamp / 86400000) + 25569;
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
