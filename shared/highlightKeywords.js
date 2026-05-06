/*
 * shared/highlightKeywords.js
 *
 * Single source of truth for the substring keywords that trigger an
 * auto-highlight when a user types in the Outreach column. Both runtimes
 * (commands/, react/) should import from here — never inline the list.
 *
 * Matching is case-insensitive substring (lowercased input .includes()
 * lowercased keyword), so each entry is one canonical lowercase form.
 */

// Phrases that indicate the student was successfully contacted /
// committed to a follow-up. Triggers a "Contacted" tag and a yellow
// row highlight when typed in the Outreach column.
export const CONTACTED_KEYWORDS = [
    'hung up',
    'hanged up',
    'promise',
    'requested',
    'up to date',
    'will catch up',
    'will come',
    'will complete',
    'will engage',
    'will pass',
    'will submit',
    'will work',
    'will be in class',
    'waiting for instructor',
    'waiting for professor',
    'waiting for teacher',
    'waiting on instructor',
    'waiting on professor',
    'waiting on teacher',
];
