/*
 * shared-utilities.js
 *
 * Constants and utilities used across the commands runtime.
 * Cross-runtime values (sheet names, batch size, hyperlink helpers,
 * findColumnIndex, column aliases) live in /shared/ and are re-exported
 * from here so existing `from './shared-utilities.js'` imports keep working.
 */
import { MASTER_LIST_SHEET, HISTORY_SHEET } from '../shared/constants.js';
import {
    STUDENT_NAME_ALIASES,
    OUTREACH_ALIASES,
    GRADE_ALIASES,
    DAYS_OUT_ALIASES,
    GRADE_BOOK_ALIASES,
    STUDENT_EMAIL_ALIASES,
    PERSONAL_EMAIL_ALIASES,
} from '../shared/columnAliases.js';
export { findColumnIndex } from '../shared/excel-helpers.js';

export const CONSTANTS = {
    // NOTE: "Student ID" and "Student Number" are treated as distinct values.
    // "Student ID" refers to SyStudentId - the school's SIS (Student Information System) ID.
    // "Student Number" (and "Student Identifier") refers to another school-based student identifier.
    // Note: SyStudentId is different from Canvas's internal Canvas ID.
    STUDENT_NAME_COLS: STUDENT_NAME_ALIASES,
    OUTREACH_COLS: OUTREACH_ALIASES,
    STUDENT_ID_COLS: ["student id", "systudentid", "id"],
    STUDENT_NUMBER_COLS: ["studentnumber", "student identifier", "student number"],
    MASTER_LIST_SHEET,
    HISTORY_SHEET,
    SETTINGS_KEY: "studentRetentionSettings", // Key for document settings
    COLUMN_MAPPINGS: {
        course: ["course"],
        courseId: ["course id"],
        courseLastAccess: ["course last access"],
        currentScore: ["current score", "grade", "course grade"],
        grade: GRADE_ALIASES,
        gradeBook: GRADE_BOOK_ALIASES,
        daysOut: DAYS_OUT_ALIASES,
        lastLda: ["lda", "last lda", "LDA", "Last LDA"],
        assigned: ["assigned", "Assigned"],
        programVersion: ["programversion", "program version"],
        courseMissingAssignments: ["course missing assignments", "missing assignments", "missingassignments"],
        courseZeroAssignments: ["course zero assignments", "zero assignments", "zeroassignments"],
        hold: ["hold", "Hold"],
        primaryPhone: ["primary phone", "phone", "phonenumber", "contact number", "Phone", "Phone Number", "Contact Number"],
        otherPhone: ["other phone", "otherphone", "Other Phone", "OtherPhone"],
        PersonalEmail: PERSONAL_EMAIL_ALIASES,
        StudentEmail: STUDENT_EMAIL_ALIASES,
        nextAssignmentDue: ["next assignment due", "Next Assignment Due", "nextassignmentdue"],
        lastCourseGrade: ["last course grade", "Last Course Grade", "lastcoursegrade"],
        expectedStartDate: ["expected start date", "start date", "expstartdate"],
        attendance: ["attendance %", "attendance%", "attendancepercent", "attendance"],
        letterGrade: ["letter grade", "lettergrade", "Letter Grade", "LetterGrade"],
        lastCourseLetterGrade: ["last course letter grade", "lastcourselettergrade", "Last Course Letter Grade", "LastCourseLetterGrade"],
        enrollGpa: ["enroll gpa", "enrollgpa", "Enroll GPA", "EnrollGPA", "Enroll Gpa"],
        courseStart: ["course start", "coursestart", "Course Start", "CourseStart"],
        courseEnd: ["course end", "courseend", "Course End", "CourseEnd"]
    }
};

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

