// Import ExcelAPI functions and expose small wrappers for use elsewhere.
import { insertRow, editRow, deleteRow } from './ExcelAPI';
import { formatTimestamp } from './Conversion';
import { Sheets } from './ColumnMapping';

// Lightweight, safe implementation that logs the outreach by inserting a row
// into the History sheet (uses Sheets.HISTORY).
// Non-blocking from callers — callers may call without awaiting.
export async function addComment(commentText, tag, createdBy = null, studentId = null, studentName = null) {
  if (!commentText) return;
  
  let userName = createdBy;
  try {
    if (!userName) {
      // prefer imported SSO helper if available
      if (typeof SSO !== 'undefined' && SSO && typeof SSO.getUserName === 'function') {
        userName = SSO.getUserName();
      }
      // fallback to localStorage keys used elsewhere in the app
      if (!userName && typeof window !== 'undefined') {
        userName = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
      }
    }
  } catch (_) {
    // ignore resolution errors
  }
  if (!userName) userName = 'Unknown';

  try {
    await insertRow(Sheets.HISTORY, {
      ID: studentId !== null && studentId !== undefined ? studentId : 0,
      Student: studentName ? String(studentName) : 'Unknown Student',
      Comment: String(commentText),
      Timestamp: formatTimestamp(new Date()),
      CreatedBy: String(userName),
      Tag: tag,
    });

    // Log who created the comment and for which student
    try {
      console.log(`${userName} added a new comment for ${studentName ? String(studentName) : 'Unknown Student'}`);
    } catch (_) {}
  } catch (err) {
    // fallback to console logging if insert fails
    try { console.error('Comment insert failed:', err); } catch (_) {}
  }
}

// Edit an existing history row (rowId can be whatever identifier ExcelAPI expects)
export async function editComment(rowId, updates) {
  // forwards to ExcelAPI.editRow
  return editRow(rowId, updates);
}

// Delete a history row
export async function deleteComment(rowId) {
  // forwards to ExcelAPI.deleteRow
  return deleteRow(rowId);
}