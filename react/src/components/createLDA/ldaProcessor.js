/*
 * Timestamp: 2026-01-26 00:00:00
 * Version: 2.22.0
 * Author: Gemini (for Victor)
 * Description: Core logic for creating LDA reports.
 * Update: Optimize formatting performance by using range-based operations instead of cell-by-cell.
 *         Consecutive cells with the same color are merged into single range operations, reducing
 *         API calls from thousands to hundreds per batch. This dramatically speeds up formatting.
 */

import { getWorkbookSettings } from '../utility/getSettings';
import { defaultColumns } from '../settings/DefaultSettings';

// Hardcoded sheet names (unless these are also settings, usually they are static)
const SHEET_NAMES = {
    MASTER_LIST: "Master List",
    HISTORY: "Student History"
};

// Batch size for chunked operations to avoid payload size limits
// Excel Add-ins have ~5MB payload limits; 500 rows is a safe batch size for data
const BATCH_SIZE = 500;

// Smaller batch size for formatting operations (colors, fonts, formulas per cell)
// These are more expensive as each cell operation adds to the request queue
// Kept small (50) to stay well under Excel's per-sync operation limits,
// especially with partial-row cell highlights generating multiple range calls per row.
const FORMAT_BATCH_SIZE = 50;

/**
 * Helper to generate the Retention Outreach Message.
 * Centralized logic for LDA tags, Missing Assignments, and Explicit DNC.
 * @param {string} sId - Student ID
 * @param {Map} ldaMap - The map of LDA data
 * @param {number} missingVal - The value from "Missing Assignments" column
 * @param {string} tableContext - 'LDA_Table' or 'Failing_Table'
 * @param {Map} dncMap - Map of ID -> Tag Text
 * @param {*} nextAssignmentDueVal - The value from "Next Assignment Due" column
 * @param {boolean} nextAssignmentDueColumnAllBlank - True if the entire column is blank
 * @param {number} daysOutVal - The student's Days Out value (used to suppress
 *   the "next assignment due" message when the student is 10+ days out)
 * @returns {string|null} - The formatted message or null
 */
function getRetentionMessage(sId, ldaMap, missingVal, tableContext, dncMap, nextAssignmentDueVal, nextAssignmentDueColumnAllBlank, daysOutVal) {
    // Priority 1: Explicit DNC (Highest Priority - Stop everything)
    if (sId && dncMap.has(sId)) {
        const dncTag = dncMap.get(sId);
        if (dncTag) {
            // Split comma-separated tags and check each individually
            const individualTags = dncTag.split(',').map(t => t.trim());
            const hasExcludableDnc = individualTags.some(tag =>
                tag.includes('dnc') && tag !== 'dnc - phone' && tag !== 'dnc - other phone'
            );
            if (hasExcludableDnc) {
                return "Do not contact";
            }
        }
    }

    // Priority 4: Zero missing assignments with a next assignment due date
    // Skipped for students 10+ days out — at that point the outreach should
    // focus on re-engagement, not on an upcoming assignment.
    if (typeof missingVal === 'number' && missingVal === 0 && nextAssignmentDueVal) {
        const isTenOrMoreDaysOut = typeof daysOutVal === 'number' && daysOutVal >= 10;
        if (!isTenOrMoreDaysOut) {
            const formattedDate = formatFriendlyDate(nextAssignmentDueVal);
            return `Student's next assignment is due ${formattedDate}.`;
        }
    }

    return null;
}
