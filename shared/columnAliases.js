/*
 * shared/columnAliases.js
 *
 * Header alias arrays for column concepts that exist in both runtimes'
 * workbooks. Each alias is lowercase; findColumnIndex (in
 * shared/excel-helpers.js) lowercases its own input, so listing each
 * unique alias once is sufficient.
 *
 * Only concepts shared by both runtimes live here. Runtime-specific
 * aliases (commands' course-* fields, React's Tags / Assignment, the
 * Student ID / Student Number distinction, etc.) stay in their
 * respective constant files.
 *
 * The aliases listed are the union of what each runtime previously
 * matched, deduped after lowercasing. Where one runtime's set was a
 * superset of the other's, the broader set wins — that runtime's
 * behavior is unchanged, the narrower one gains a few extra matches.
 */

export const STUDENT_NAME_ALIASES = ['studentname', 'student name', 'student'];

export const OUTREACH_ALIASES = ['outreach'];

export const GRADE_ALIASES = ['grade', 'course grade', 'grades'];

export const DAYS_OUT_ALIASES = ['days out', 'daysout'];

export const GRADE_BOOK_ALIASES = [
    'grade book',
    'gradebook',
    'gradebooklink',
    'gradelink',
];

export const STUDENT_EMAIL_ALIASES = [
    'student email',
    'school email',
    'email',
    'studentemail',
];

export const PERSONAL_EMAIL_ALIASES = [
    'personal email',
    'otheremail',
    'personalemail',
];
