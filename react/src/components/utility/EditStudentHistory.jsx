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

	// date part YYMMDD (use provided Timestamp or today) — changed to two-digit year
	let datePart;
	try {
		let d = Timestamp ? new Date(Timestamp) : new Date();
		if (isNaN(d.getTime())) d = new Date();
		const yyyy = d.getFullYear();
		const yy = String(yyyy).slice(-2).padStart(2, '0'); // two-digit year
		const mm = String(d.getMonth() + 1).padStart(2, '0');
		const dd = String(d.getDate()).padStart(2, '0');
		datePart = `${yy}${mm}${dd}`;
	} catch (_) {
		const d = new Date();
		const yyyy = d.getFullYear();
		const yy = String(yyyy).slice(-2).padStart(2, '0');
		datePart = `${yy}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
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

	// final ID: studentLast4 + YYMMDD + tagLast4
  console.log(Tag)
  console.log(tagLast4)
  console.log(tagDigits)
	return `${studentLast4}${datePart}${tagLast4}`;
}

// Lightweight, safe implementation that logs the outreach by inserting a row
// into the History sheet (uses Sheets.HISTORY).
// Non-blocking from callers — callers may call without awaiting.
export async function addComment(commentText, tag, createdBy = null, studentId = null, studentName = null) {
  if (!commentText) return;
  
  const userName = checkSSO(createdBy);

  try {
    // Check if an Outreach comment for this student exists today
    const todaysOutreach = await checkTodaysOutreach(studentId);
    const ts = formatTimestamp(new Date());

    if (!todaysOutreach) {
      // No outreach today -> insert a new row
      const result = await insertRow(Sheets.HISTORY, {
        ID: studentId !== null && studentId !== undefined ? studentId : 0,
        Student: studentName ? String(studentName) : 'Unknown Student',
        Comment: String(commentText),
        Timestamp: ts,
        CreatedBy: String(userName),
        Tag: tag,
        commentID: generateCommentID(studentId, new Date(), tag)
      });
      try { console.log(`${userName} inserted a new comment for ${studentName ? String(studentName) : 'Unknown Student'}`); } catch (_) {}
      return result;
    } else {
      // Outreach exists today -> edit the existing Outreach row (use helper editComment)
      const outreachCommentID = generateCommentID(studentId, new Date(), 'Outreach');
      const updates = {
        Comment: String(commentText),
        Timestamp: ts,
        CreatedBy: String(userName),
        Tag: tag
      };
      const result = await editComment(outreachCommentID, updates);
      try { console.log(`${userName} edited today's Outreach comment for ${studentName ? String(studentName) : 'Unknown Student'}`); } catch (_) {}
      return result;
    }
  } catch (err) {
    try { console.error('Comment insert/edit failed:', err); } catch (_) {}
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
export async function checkTodaysOutreach(studentId = null) {
  try {
    const CommentID = generateCommentID(studentId, new Date(), 'Outreach');
    const res = await checkRow(Sheets.HISTORY, 'commentID', CommentID);
    // Return true only if checkRow explicitly returned true, otherwise false
    return res === true;
  } catch (err) {
    try { console.error('checkTodaysOutreach failed:', err); } catch (_) {}
    return false;
  }
}