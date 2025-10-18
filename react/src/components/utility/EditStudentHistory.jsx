// Import ExcelAPI functions and expose small wrappers for use elsewhere.
import { insertRow, editRow, deleteRow } from './ExcelAPI';
import { formatTimestamp } from './Conversion';
import { Sheets } from './ColumnMapping';

// New: small helper to resolve SSO/localStorage username with safe fallbacks
function checkSSO(provided) {
  let user = provided;
  try {
    if (!user) {
      if (typeof SSO !== 'undefined' && SSO && typeof SSO.getUserName === 'function') {
        user = SSO.getUserName();
      }
      if (!user && typeof window !== 'undefined') {
        user = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
      }
    }
  } catch (_) {
    // ignore resolution errors
  }
  return user || 'Unknown';
}

// New: generate numeric comment IDs (prefers a localStorage counter, falls back to timestamp)
function generateCommentID() {
  try {
    // prefer secure RNG when available
    if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
      const buf = new Uint32Array(3);
      crypto.getRandomValues(buf);
      // concatenate and extract digits, ensure at least 10 chars
      let digits = Array.from(buf).map(n => String(Math.abs(n))).join('').replace(/\D/g, '').slice(0, 10);
      while (digits.length < 10) {
        digits += String(Math.floor(Math.random() * 10));
      }
      return digits;
    }
  } catch (_) {
    // ignore and fallback
  }

  // fallback: Math.random-based 10-digit string
  const n = Math.floor(Math.random() * 1e10);
  return String(n).padStart(10, '0');
}

// Lightweight, safe implementation that logs the outreach by inserting a row
// into the History sheet (uses Sheets.HISTORY).
// Non-blocking from callers — callers may call without awaiting.
export async function addComment(commentText, tag, createdBy = null, studentId = null, studentName = null) {
  if (!commentText) return;
  
  const userName = checkSSO(createdBy);

  try {
    await insertRow(Sheets.HISTORY, {
      ID: studentId !== null && studentId !== undefined ? studentId : 0,
      Student: studentName ? String(studentName) : 'Unknown Student',
      Comment: String(commentText),
      Timestamp: formatTimestamp(new Date()),
      CreatedBy: String(userName),
      Tag: tag,
      commentID: generateCommentID() // numeric ID
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

// Delete a history row by commentID (numeric or string)
export async function deleteComment(commentID, createdBy = null, studentId = null, studentName = null) {
  if (commentID === null || commentID === undefined) return;

  const userName = checkSSO(createdBy);

  try {
    // attempt to delete the row — forward the commentID to ExcelAPI.deleteRow
    const result = await deleteRow(commentID);

    // Log who performed the deletion (no insertRow audit)
    try {
      console.log(`${userName} deleted comment ${String(commentID)} for ${studentName ? String(studentName) : 'Unknown Student'}`);
    } catch (_) {}

    return result;
  } catch (err) {
    // deletion failed — log and rethrow
    try { console.error('Comment delete failed:', err); } catch (_) {}
    throw err;
  }
}