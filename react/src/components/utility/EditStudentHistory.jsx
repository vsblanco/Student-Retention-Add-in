// Import ExcelAPI functions and expose small wrappers for use elsewhere.
import { insertRow, editRow, deleteRow, checkRow } from './ExcelAPI';
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
function generateCommentID(StudentID = null, Timestamp = null, Tag = null) {
	// last 4 numeric digits of StudentID (pad to 4)
	let studentDigits = '';
	try {
		studentDigits = StudentID !== null && StudentID !== undefined ? String(StudentID).replace(/\D/g, '') : '';
	} catch (_) {
		studentDigits = '';
	}
	let studentLast4 = studentDigits.slice(-4);
	if (!studentLast4) studentLast4 = '0000';
	else if (studentLast4.length < 4) studentLast4 = studentLast4.padStart(4, '0');

	// date part YYYYMMDD (use provided Timestamp or today)
	let datePart;
	try {
		let d = Timestamp ? new Date(Timestamp) : new Date();
		if (isNaN(d.getTime())) d = new Date();
		const yyyy = d.getFullYear();
		const mm = String(d.getMonth() + 1).padStart(2, '0');
		const dd = String(d.getDate()).padStart(2, '0');
		datePart = `${yyyy}${mm}${dd}`;
	} catch (_) {
		const d = new Date();
		datePart = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
	}

	// last 4 digits derived from Tag (map to char codes then take last 4 digits, pad to 4)
	let tagDigits = '';
	try {
		const t = Tag !== null && Tag !== undefined ? String(Tag) : '';
		tagDigits = t.split('').map(c => String(c.charCodeAt(0))).join('').replace(/\D/g, '');
	} catch (_) {
		tagDigits = '';
	}
	let tagLast4 = tagDigits.slice(-4);
	if (!tagLast4) tagLast4 = '0000';
	else if (tagLast4.length < 4) tagLast4 = tagLast4.padStart(4, '0');

	// final ID: studentLast4 + YYYYMMDD + tagLast4
	return `${studentLast4}${datePart}${tagLast4}`;
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
      commentID: generateCommentID(studentId, new Date(), tag) // studentLast4 + YYYYMMDD + tagLast4
    });

    // Log who created the comment and for which student
    try {
      console.log(await checkTodaysOutreach());
      console.log(`${userName} added a new comment for ${studentName ? String(studentName) : 'Unknown Student'}`);
    } catch (_) {}
  } catch (err) {
    // fallback to console logging if insert fails
    try { console.error('Comment insert failed:', err); } catch (_) {}
  }
}

// Edit an existing history row (rowId can be whatever identifier ExcelAPI expects)
export async function editComment(commentid, updates) {
  // forwards to ExcelAPI.editRow
  return editRow(Sheets.HISTORY, "commentID", commentid, updates);
}

// Delete a history row by commentID (numeric or string)
export async function deleteComment(commentID, createdBy = null) {
  if (commentID === null || commentID === undefined) return;

  const userName = checkSSO(createdBy);

  try {
    // attempt to delete the row — forward the commentID to ExcelAPI.deleteRow
    const result = await deleteRow(Sheets.HISTORY, "commentID", commentID);

    // Log who performed the deletion (no insertRow audit)
    try {
      console.log(`${userName} deleted comment ${String(commentID)} : 'Unknown Student'}`);
    } catch (_) {}

    return result;
  } catch (err) {
    // deletion failed — log and rethrow
    try { console.error('Comment delete failed:', err); } catch (_) {}
    throw err;
  }
}

// New: check whether there's an "Outreach" comment with a timestamp that matches today.
// Returns true if at least one Outreach row has a Timestamp on the current date.
export async function checkTodaysOutreach() {
  try {
    const res = await checkRow(Sheets.HISTORY, 'Tag', 'Outreach');

    if (!res) return false;

    // Normalize to array
    const rows = Array.isArray(res) ? res : [res];

    const today = new Date();
    const isoToday = today.toISOString().slice(0, 10); // YYYY-MM-DD

    const parseAndMatch = (ts) => {
      if (!ts) return false;

      // If it's already a Date or parseable ISO, use it
      try {
        const d = new Date(ts);
        if (!isNaN(d.getTime())) {
          return d.getFullYear() === today.getFullYear()
            && d.getMonth() === today.getMonth()
            && d.getDate() === today.getDate();
        }
      } catch (_) {}

      const s = String(ts);

      // Try to find an ISO date substring
      const iso = s.match(/(\d{4}-\d{2}-\d{2})/);
      if (iso && iso[1] === isoToday) return true;

      // Try common US format MM/DD/YYYY or M/D/YYYY
      const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) {
        const mm = parseInt(m[1], 10) - 1;
        const dd = parseInt(m[2], 10);
        const yyyy = parseInt(m[3], 10);
        return yyyy === today.getFullYear() && mm === today.getMonth() && dd === today.getDate();
      }

      return false;
    };

    for (const r of rows) {
      const ts = r && typeof r === 'object' ? (r.Timestamp || r.timestamp || r.Time || r.time) : r;
      if (parseAndMatch(ts)) return true;
    }

    return false;
  } catch (err) {
    try { console.error('checkTodaysOutreach failed:', err); } catch (_) {}
    return false;
  }
}