import { describe, it, expect } from 'vitest';
import { findColumnIndex } from '../../shared/excel-helpers.js';
import {
    STUDENT_NAME_ALIASES,
    OUTREACH_ALIASES,
    GRADE_ALIASES,
    DAYS_OUT_ALIASES,
    GRADE_BOOK_ALIASES,
    STUDENT_EMAIL_ALIASES,
    PERSONAL_EMAIL_ALIASES,
} from '../../shared/columnAliases.js';

// Aliases each runtime PREVIOUSLY matched (before the partial merge into shared/).
// Every entry below must still be reachable via findColumnIndex against the shared list.
// If this test fails, the merge dropped an alias that one of the runtimes was relying on.
const PREVIOUSLY_MATCHED = {
    STUDENT_NAME_ALIASES: {
        commands: ['studentname', 'student name', 'student'],
        react: ['studentname', 'student name'],
    },
    OUTREACH_ALIASES: {
        commands: ['outreach'],
        react: ['outreach'],
    },
    GRADE_ALIASES: {
        commands: ['grade', 'course grade', 'grades'],
        react: ['grade', 'course grade', 'grades'],
    },
    DAYS_OUT_ALIASES: {
        commands: ['days out'],
        react: ['days out', 'daysout'],
    },
    GRADE_BOOK_ALIASES: {
        commands: ['grade book', 'gradebook', 'gradebooklink', 'gradelink'],
        react: ['grade book', 'gradebook'],
    },
    STUDENT_EMAIL_ALIASES: {
        commands: ['student email', 'school email', 'email'],
        react: ['student email', 'school email', 'email', 'studentemail'],
    },
    PERSONAL_EMAIL_ALIASES: {
        commands: ['personal email', 'otheremail'],
        react: ['personal email', 'otheremail', 'personalemail'],
    },
};

const ALIAS_LISTS = {
    STUDENT_NAME_ALIASES,
    OUTREACH_ALIASES,
    GRADE_ALIASES,
    DAYS_OUT_ALIASES,
    GRADE_BOOK_ALIASES,
    STUDENT_EMAIL_ALIASES,
    PERSONAL_EMAIL_ALIASES,
};

describe('shared/columnAliases shape', () => {
    it.each(Object.entries(ALIAS_LISTS))(
        '%s is a non-empty array of unique lowercase strings',
        (name, list) => {
            expect(Array.isArray(list)).toBe(true);
            expect(list.length).toBeGreaterThan(0);
            for (const alias of list) {
                expect(typeof alias).toBe('string');
                expect(alias).toBe(alias.toLowerCase());
            }
            expect(new Set(list).size).toBe(list.length); // no duplicates
        }
    );
});

describe('shared/columnAliases preserves previous matches', () => {
    for (const [listName, byRuntime] of Object.entries(PREVIOUSLY_MATCHED)) {
        const list = ALIAS_LISTS[listName];

        for (const [runtime, aliases] of Object.entries(byRuntime)) {
            for (const alias of aliases) {
                it(`${listName} still matches ${runtime}'s "${alias}"`, () => {
                    // Simulate the lookup the way callers do: lowercased headers
                    // (just the alias as a header), searched via findColumnIndex.
                    const headers = [alias.toLowerCase()];
                    expect(findColumnIndex(headers, list)).toBe(0);
                });
            }
        }
    }
});
