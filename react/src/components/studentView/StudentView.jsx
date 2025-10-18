import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
import StudentHeader from './StudentHeader.jsx';
import StudentAssignments from './StudentAssignments.jsx';
import { formatName } from '../utility/Conversion.jsx';
import { onSelectionChanged } from '../utility/ExcelAPI.jsx';
import SSO from '../utility/SSO.jsx';
import { getCanonicalColIdx } from '../utility/CanonicalMap.jsx';
import './StudentView.css';
import { loadCache } from '../utility/Cache.jsx';
import { insertRow } from '../utility/ExcelAPI.jsx';
import { Sheets } from '../utility/ColumnMapping.jsx';

// Replace the previous OUTREACH_TRIGGERS array with a normalized, deduplicated, sorted list
const OUTREACH_TRIGGERS = [
  "hung up",
  "hanged up",
  "promise",
  "requested",
  "up to date",
  "will catch up",
  "will come",
  "will complete",
  "will engage",
  "will pass",
  "will submit",
  "will work",
  "will be in class",
  "waiting for instructor",
  "waiting for professor",
  "waiting for teacher",
  "waiting on instructor",
  "waiting on professor",
  "waiting on teacher"
];

// Precompute lowercase triggers for faster, case-insensitive checks
const OUTREACH_TRIGGERS_LOWER = OUTREACH_TRIGGERS.map(t => t.toLowerCase());

const isOutreachTrigger = (text) => {
  if (!text || typeof text !== 'string') return false;
  const lower = text.toLowerCase();
  return OUTREACH_TRIGGERS_LOWER.some(trigger => lower.includes(trigger));
};

// helper: format date as "MM/DD/YY HH:MM AM/PM"
function formatTimestamp(date = new Date()) {
  try {
    const d = new Date(date);
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const yy = String(d.getFullYear() % 100).padStart(2, '0');

    let hours = d.getHours();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    if (hours === 0) hours = 12;
    const hh = String(hours).padStart(2, '0');
    const mins = String(d.getMinutes()).padStart(2, '0');

    return `${mm}/${dd}/${yy} ${hh}:${mins} ${ampm}`;
  } catch (_) {
    return String(date);
  }
}

// Lightweight, safe implementation that logs the outreach by inserting a row
// into the History sheet (uses Sheets.HISTORY).
// Non-blocking from callers — callers may call without awaiting.
async function addComment(commentText, tag, createdBy = 'Unknown', studentId = null, studentName = null) {
  if (!commentText) return;
  try {
    await insertRow(Sheets.HISTORY, {
      ID: studentId !== null && studentId !== undefined ? studentId : 0,
      Student: studentName ? String(studentName) : 'Unknown Student',
      Comment: String(commentText),
      Timestamp: formatTimestamp(new Date()),
      CreatedBy: String(createdBy),
      Tag: tag,
    });
  } catch (err) {
    // fallback to console logging if insert fails
    try { console.error('Comment insert failed:', err); } catch (_) {}
  }
}

async function highlightRow(rowIndex, startCol, colCount, color = 'yellow') {
  if (typeof window.Excel === "undefined") return;
  if (typeof rowIndex !== 'number' || typeof startCol !== 'number' || typeof colCount !== 'number') return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const highlightRange = sheet.getRangeByIndexes(rowIndex, startCol, 1, colCount);
      highlightRange.format.fill.color = color;
      await context.sync();
    });
  } catch (_) {
    // swallow errors
  }
}

function StudentView() {
  const [activeStudent, setActiveStudent] = useState(null);
  const [sheetData, setSheetData] = useState({ status: 'loading', data: null, message: 'Loading student data...' });
  const [activeTab, setActiveTab] = useState('details');
  const isInitialLoad = useRef(true);
  const [headers, setHeaders] = useState([]);
  const [assignmentsMap, setAssignmentsMap] = useState({});
  const [userName, setUserName] = useState(null);

  // use a ref to keep the current session user available to async handlers
  const sessionCommentUserRef = useRef(null);

  // Keep last source used when setting activeStudent so we can log once in an effect
  const lastSetSourceRef = useRef(null);
  function setActiveStudentWithLog(student, source = 'unknown') {
    lastSetSourceRef.current = source;
    setActiveStudent(student);
  }

  // Log activeStudent once when it actually changes (avoids render-time and StrictMode duplicates)
  useEffect(() => {
    try {
      if (activeStudent) {
        console.log(`activeStudent (source: ${lastSetSourceRef.current || 'unknown'}):`, activeStudent);
      } else {
        console.log(`activeStudent cleared (source: ${lastSetSourceRef.current || 'unknown'})`);
      }
    } catch (_) {}
    lastSetSourceRef.current = null;
  }, [activeStudent]);

  // Initialize userName from cache/SSO on mount
  useEffect(() => {
    try {
      const cached = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
      if (cached) {
        setUserName(cached);
        sessionCommentUserRef.current = cached;
        return;
      }
      if (window.SSO && typeof window.SSO.getUserName === 'function') {
        const n = window.SSO.getUserName();
        if (n) {
          setUserName(n);
          sessionCommentUserRef.current = n;
        }
      }
    } catch (_) { /* ignore */ }
  }, []);

  // Persist userName and keep sessionCommentUserRef in sync
  useEffect(() => {
    if (!userName) return;
    try { window.localStorage.setItem('ssoUserName', userName); } catch (_) {}
    sessionCommentUserRef.current = userName;
  }, [userName]);

  // Error handler
  const errorHandler = (error) => {
    console.error("onWorksheetChanged error:", error);
  };

  const isHandlerRunning = useRef(false);

  // Use a stable handler that reads sessionCommentUserRef to avoid stale closures
  async function onWorksheetChanged(eventArgs) {
    if (typeof window.Excel === "undefined") return;
    if (!eventArgs || !eventArgs.address) return;
    if (isHandlerRunning.current) return;
    isHandlerRunning.current = true;
    try {
      await Excel.run(async (context) => {
        if (eventArgs.source !== Excel.EventSource.local || (eventArgs.changeType !== "CellEdited" && eventArgs.changeType !== "RangeEdited")) {
          return;
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const changedRange = sheet.getRange(eventArgs.address);
        changedRange.load("rowIndex, columnIndex, rowCount, columnCount, values");
        const headerRange = sheet.getRange("1:1").getUsedRange(true);
        headerRange.load("values, columnCount");
        await context.sync();

        if (!headerRange || !headerRange.values || !headerRange.values[0]) return;
        const headersRow = headerRange.values[0];
        const outreachColIndex = getCanonicalColIdx(headersRow, 'Outreach');

        if (outreachColIndex === -1 || changedRange.rowIndex === 0) return;

        const outreachColumnAffected = changedRange.columnIndex <= outreachColIndex && (changedRange.columnIndex + changedRange.columnCount - 1) >= outreachColIndex;
        if (!outreachColumnAffected) return;

        const studentInfoRange = sheet.getRangeByIndexes(
          changedRange.rowIndex, 0,
          changedRange.rowCount, headerRange.columnCount
        );
        studentInfoRange.load("values");
        await context.sync();
        const allRowValues = studentInfoRange.values || [];

        const studentIdColIndex = getCanonicalColIdx(headersRow, 'ID');
        const studentNameColIndex = getCanonicalColIdx(headersRow, 'StudentName');

        if (studentIdColIndex === -1 || studentNameColIndex === -1) return;

        for (let i = 0; i < changedRange.rowCount; i++) {
          const currentRow = allRowValues[i] || [];
          const newValueRaw = currentRow[outreachColIndex];
          const newValue = (newValueRaw !== undefined && newValueRaw !== null) ? String(newValueRaw).trim() : "";
          if (!newValue) continue;

          const studentId = currentRow[studentIdColIndex];
          const studentName = currentRow[studentNameColIndex];
          const rowIndex = changedRange.rowIndex + i;

          if (studentId && studentName) {
            const effectiveUser = sessionCommentUserRef.current || userName || window.localStorage.getItem('ssoUserName') || 'Unknown';
            // If the comment text matches a trigger, tag as 'Contacted' and highlight the row.
            if (isOutreachTrigger(newValue)) {
              // await the comment insert, then highlight, then refresh cache
              try {
                await addComment(newValue, 'Contacted', effectiveUser, studentId, studentName);
              } catch (_) { /* ignore insert errors */ }
              try {
                await highlightRow(
                  rowIndex,
                  Math.min(studentNameColIndex, outreachColIndex),
                  Math.abs(studentNameColIndex - outreachColIndex) + 1,
                  'yellow'
                );
              } catch (_) { /* ignore highlight errors */ }
              // refresh the cached data to reflect the new comment
              await refreshCache();
            } else {
              // await insert then refresh so UI updates
              try {
                await addComment(newValue, 'Outreach', effectiveUser, studentId, studentName);
                await refreshCache();
              } catch (_) { /* ignore errors */ }
            }
          }
        }
      });
    } catch (error) {
      errorHandler(error);
    } finally {
      isHandlerRunning.current = false;
    }
  }

  // Add: refreshCache updates sheet data and ensures the current active student object
  // is replaced with the fresh one from the cache (causes StudentHistory to rerender).
  async function refreshCache() {
    try {
      const res = await loadCache();
      setSheetData({ status: res.status || 'success', data: res.data || {}, message: res.message || '' });
      setHeaders(res.headers || []);
      setAssignmentsMap(res.assignmentsMap || {});

      // If we have an active student, try to find the fresh object in the new cache and re-set it.
      if (res.status === 'success' && activeStudent) {
        const currentId = activeStudent.ID ?? activeStudent.Id ?? activeStudent.id;
        if (currentId !== undefined && currentId !== null) {
          const dataObj = res.data || {};
          // search through values (keys may be row indexes)
          const fresh = Object.values(dataObj).find(s => {
            if (!s) return false;
            return (s.ID == currentId || s.Id == currentId || s.id == currentId);
          });
          if (fresh) {
            setActiveStudentWithLog(fresh, 'refreshCache');
          }
        }
      }
    } catch (err) {
      // swallow errors silently — existing code prefers non-blocking behavior
    }
  }

  // Effect: Load sheet cache (Excel or test-mode) once
  useEffect(() => {
    let mounted = true;
    const run = async () => {
      setSheetData({ status: 'loading', data: null, message: 'Loading student data...' });
      try {
        const res = await loadCache();
        if (!mounted) return;
        setSheetData({ status: res.status || 'success', data: res.data || {}, message: res.message || '' });
        setHeaders(res.headers || []);
        setAssignmentsMap(res.assignmentsMap || {});
        if (res.status === 'success' && (!window.Excel || Object.keys(res.data || {}).length === 1)) {
          const firstKey = Object.keys(res.data || [])[0];
          if (firstKey)
            setActiveStudentWithLog(res.data[firstKey], 'initialLoad');
        }
      } catch (err) {
        if (!mounted) return;
        setSheetData({ status: 'error', data: null, message: 'An error occurred while loading the data. Please try again.' });
      }
    };
    run();
    return () => { mounted = false; };
  }, []);

  // Keep selection handler ref for cleanup
  const selectionHandlerRef = useRef(null);
  const worksheetHandlerRef = useRef(null);
  const isHandlerAttached = useRef(false);

  // Effect: Handle selection changes from Excel.
  useEffect(() => {
    if (typeof window.Excel === "undefined") return;
    if (sheetData.status !== 'success') return;

    let eventHandlerObj = null;

    const syncUIToSheetSelection = async () => {
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load(["rowIndex", "columnIndex"]);
          await context.sync();
          const rowIndex = range.rowIndex;
          const colIndex = range.columnIndex;
          setActiveStudentWithLog(sheetData.data[rowIndex] || null, 'selection');

          const outreachIdx = getCanonicalColIdx(headers, 'Outreach');
          const missingAssignmentsIdx = getCanonicalColIdx(headers, 'Missing Assignments');
          const phoneIdx = getCanonicalColIdx(headers, 'Phone');
          const otherPhoneIdx = getCanonicalColIdx(headers, 'OtherPhone');
          if (
            (outreachIdx !== -1 && colIndex === outreachIdx) ||
            (phoneIdx !== -1 && colIndex === phoneIdx) ||
            (otherPhoneIdx !== -1 && colIndex === otherPhoneIdx)
          ) {
            setActiveTab('history');
          } else if (missingAssignmentsIdx !== -1 && colIndex === missingAssignmentsIdx) {
            setActiveTab('assignments');
          } else {
            setActiveTab('details');
          }
        });
      } catch (error) {
        if (error.code !== "GeneralException") {
          console.error("Error in syncUIToSheetSelection:", error);
        }
      }
    };

    onSelectionChanged(() => {
      syncUIToSheetSelection();
    }).then(obj => {
      eventHandlerObj = obj;
      selectionHandlerRef.current = obj;
      if (isInitialLoad.current) {
        syncUIToSheetSelection();
        isInitialLoad.current = false;
      }
    });

    // Attach worksheet changed handler once
    if (!isHandlerAttached.current) {
      Excel.run(async (context) => {
        try {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const handler = sheet.onChanged.add(onWorksheetChanged);
          worksheetHandlerRef.current = handler;
          isHandlerAttached.current = true;
        } catch (err) {
          // ignore attach errors
        }
      }).catch(() => {});
    }

    return () => {
      if (eventHandlerObj && eventHandlerObj.remove) {
        try { eventHandlerObj.remove(); } catch (_) {}
      }
      selectionHandlerRef.current = null;
      if (isHandlerAttached.current && worksheetHandlerRef.current && worksheetHandlerRef.current.remove) {
        try {
          worksheetHandlerRef.current.remove();
        } catch (_) {}
        isHandlerAttached.current = false;
      }
      worksheetHandlerRef.current = null;
    };
  }, [sheetData, headers]);

  // Detect test mode (browser, not Excel)
  const isTestMode = typeof window.Excel === "undefined";

  // Always show SSO first until userName is set
  if (!userName) {
    return (
      <div className="studentview-outer">
        <SSO onNameSelect={setUserName} />
      </div>
    );
  }

  // Centralized rendering logic based on sheetData.status
  if (sheetData.status !== 'success') {
    return (
      <div className="studentview-outer">
        <div className="studentview-message"><p>{sheetData.message}</p></div>
      </div>
    );
  }

  if (!activeStudent) {
    return (
      <div className="studentview-outer">
        <div className="studentview-message">
          <p>Select a cell in a student's row to view their details.</p>
        </div>
      </div>
    );
  }
   
  return (
    <div
      className={isTestMode ? "studentview-outer testmode" : "studentview-outer"}
      style={{ userSelect: "none" }}
    >
      <StudentHeader student={activeStudent} />
      <div className="studentview-tabs">
        <div
          className={`studentview-tab${activeTab === 'details' ? ' active' : ''}`}
          onClick={() => setActiveTab('details')}
        >
          Details
        </div>
        <div
          className={`studentview-tab${activeTab === 'history' ? ' active' : ''}`}
          onClick={() => setActiveTab('history')}
        >
          History
        </div>
        <div
          className={`studentview-tab${activeTab === 'assignments' ? ' active' : ''}`}
          onClick={() => setActiveTab('assignments')}
        >
          Assignments
        </div>
      </div>
      <div className="studentview-container">
        {activeTab === 'details' && <StudentDetails student={activeStudent} />}
        {activeTab === 'history' && (
          <StudentHistory
            history={activeStudent && (activeStudent.History || activeStudent.history || [])}
            student={activeStudent}
          />
        )}
        {activeTab === 'assignments' && (
          <StudentAssignments student={activeStudent} />
        )}
      </div>
    </div>
  );
}

export default StudentView;