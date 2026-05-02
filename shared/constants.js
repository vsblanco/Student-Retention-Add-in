/*
 * shared/constants.js
 *
 * Constants shared by both runtimes (React task pane and the commands runtime).
 * Anything that previously had divergent local copies in commands/ and react/
 * lives here.
 */

// Worksheet names (the actual sheet name in the workbook — not UI labels).
// If you're rendering display text in the UI, hardcode the string instead;
// these constants are for `worksheets.getItem(...)`-style lookups so a rename
// only needs to happen in one place.
export const MASTER_LIST_SHEET = 'Master List';
export const HISTORY_SHEET = 'Student History';

// Batch size for chunked Excel read/write operations to stay under
// Office's ~5MB payload limit. 500 rows is safe for typical column counts.
export const BATCH_SIZE = 500;
