import { describe, it, expect } from 'vitest';
import { findColumnIndex, normalizeHeader } from '../../shared/excel-helpers.js';
import * as Aliases from '../../shared/columnAliases.js';

// Aliases each runtime PREVIOUSLY matched (before the consolidation into shared/).
// Every entry below must still be reachable via findColumnIndex against the
// shared list. If this test fails, the consolidation dropped an alias that
// one of the runtimes was relying on.
//
// For "merged" entries (assigned, missing-assignments, etc.) where commands'
// behavior was deliberately broadened to match React's, both runtimes' lists
// are checked.
const PREVIOUSLY_MATCHED = {
    STUDENT_NAME_ALIASES: ['studentname', 'student name', 'student'],
    STUDENT_ID_ALIASES: ['student id', 'systudentid', 'id'],
    STUDENT_NUMBER_ALIASES: ['student number', 'student identifier', 'studentnumber'],
    OUTREACH_ALIASES: ['outreach'],
    HOLD_ALIASES: ['hold'],
    ASSIGNED_ALIASES: ['assigned', 'advisor', 'counselor', 'assigned to'],
    GRADE_ALIASES: ['grade', 'course grade', 'grades'],
    CURRENT_SCORE_ALIASES: ['current score', 'grade', 'course grade'],
    LAST_COURSE_GRADE_ALIASES: ['last course grade', 'lastcoursegrade'],
    LETTER_GRADE_ALIASES: ['letter grade', 'lettergrade'],
    LAST_COURSE_LETTER_GRADE_ALIASES: ['last course letter grade', 'lastcourselettergrade'],
    ENROLL_GPA_ALIASES: ['enroll gpa', 'enrollgpa'],
    STUDENT_EMAIL_ALIASES: ['student email', 'school email', 'email', 'studentemail'],
    PERSONAL_EMAIL_ALIASES: ['personal email', 'otheremail', 'personalemail'],
    PRIMARY_PHONE_ALIASES: ['primary phone', 'phone', 'phonenumber', 'contact number'],
    OTHER_PHONE_ALIASES: ['other phone', 'otherphone'],
    COURSE_ALIASES: ['course'],
    COURSE_ID_ALIASES: ['course id'],
    COURSE_LAST_ACCESS_ALIASES: ['course last access'],
    COURSE_START_ALIASES: ['course start', 'coursestart'],
    COURSE_END_ALIASES: ['course end', 'courseend'],
    PROGRAM_VERSION_ALIASES: ['programversion', 'program version'],
    DAYS_OUT_ALIASES: ['days out', 'daysout'],
    LAST_LDA_ALIASES: ['lda', 'last lda'],
    ATTENDANCE_ALIASES: ['attendance %', 'attendance%', 'attendancepercent', 'attendance'],
    EXPECTED_START_DATE_ALIASES: ['expected start date', 'start date', 'expstartdate'],
    NEXT_ASSIGNMENT_DUE_ALIASES: ['next assignment due', 'nextassignmentdue'],
    GRADE_BOOK_ALIASES: ['grade book', 'gradebook', 'gradebooklink', 'gradelink'],
    ASSIGNMENT_ALIASES: ['assignment'],
    MISSING_ASSIGNMENTS_ALIASES: ['course missing assignments', 'missing assignments', 'missingassignments', 'missing'],
    COURSE_ZERO_ASSIGNMENTS_ALIASES: ['course zero assignments', 'zero assignments', 'zeroassignments'],
    TAGS_ALIASES: ['tag', 'tags', 'history'],
};

describe('shared/columnAliases shape', () => {
    it.each(Object.entries(Aliases))(
        '%s is a non-empty array of unique normalized strings',
        (name, list) => {
            expect(Array.isArray(list)).toBe(true);
            expect(list.length).toBeGreaterThan(0);
            for (const alias of list) {
                expect(typeof alias).toBe('string');
                // After normalizeHeader, the entry should match itself —
                // i.e., entries are already in canonical form (no need
                // to enumerate case/whitespace variants).
                expect(normalizeHeader(alias)).toBe(alias.replace(/\s+/g, '').toLowerCase());
            }
            // Each entry must produce a unique normalized form
            const normalized = list.map(normalizeHeader);
            expect(new Set(normalized).size).toBe(list.length);
        }
    );
});

describe('shared/columnAliases preserves previous matches', () => {
    for (const [listName, previousAliases] of Object.entries(PREVIOUSLY_MATCHED)) {
        const list = Aliases[listName];
        if (!list) {
            it(`${listName} export exists`, () => {
                throw new Error(`${listName} not exported from shared/columnAliases.js`);
            });
            continue;
        }

        for (const alias of previousAliases) {
            it(`${listName} still matches "${alias}"`, () => {
                // Realistic call: header passed in is the (raw) alias, normalized
                // by the caller, then findColumnIndex normalizes the alias list.
                const headers = [normalizeHeader(alias)];
                expect(findColumnIndex(headers, list)).toBe(0);
            });
        }
    }
});

describe('shared/columnAliases case + whitespace insensitivity', () => {
    // Spot-check that a single concept matches case/whitespace variants
    // without the alias list having to enumerate them.
    it('GRADE_BOOK_ALIASES matches "Grade Book", "GradeBook", "GRADE BOOK"', () => {
        for (const variant of ['Grade Book', 'GradeBook', 'GRADE BOOK', '  grade  book  ']) {
            const headers = [normalizeHeader(variant)];
            expect(findColumnIndex(headers, Aliases.GRADE_BOOK_ALIASES)).toBe(0);
        }
    });

    it('STUDENT_NAME_ALIASES matches "Student Name", "STUDENT NAME", "studentname"', () => {
        for (const variant of ['Student Name', 'STUDENT NAME', 'studentname', '  student name  ']) {
            const headers = [normalizeHeader(variant)];
            expect(findColumnIndex(headers, Aliases.STUDENT_NAME_ALIASES)).toBe(0);
        }
    });
});
