import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './Tabs/Details.jsx';
import StudentHistory from './Tabs/History.jsx';
import StudentHeader from './Parts/Header.jsx';
import StudentAssignments from './Tabs/Assignments.jsx';
import { addComment } from '../utility/EditStudentHistory.jsx';
import { onSelectionChanged, highlightRow, loadSheet } from '../utility/ExcelAPI.jsx';
import SSO from '../utility/SSO.jsx';
import { getCanonicalColIdx } from '../utility/CanonicalMap.jsx';
import './Styling/StudentView.css';
import { loadCache } from '../utility/Cache.jsx';

// Outreach trigger phrases (case-insensitive substring match)
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

function StudentView() {
  const [activeStudent, setActiveStudent] = useState(null);
  const [sheetData, setSheetData] = useState({ status: 'loading', data: null, message: 'Loading student data...' });
  const [activeTab, setActiveTab] = useState('details');
  const isInitialLoad = useRef(true);
  const [headers, setHeaders] = useState([]);
  const [assignmentsMap, setAssignmentsMap] = useState({});
  // renamed for clarity: currentUserName holds the logged-in user name
  const [currentUserName, setCurrentUserName] = useState(null);

  // use a ref to keep the current session user available to async handlers
  // renamed for clarity: sessionUserRef holds the in-session user fallback for async handlers
  const sessionUserRef = useRef(null);

  // Keep last source used when setting activeStudent so we can log once in an effect
  // renamed for clarity
  const lastSelectionSourceRef = useRef(null);
  function setSelectedStudentWithSource(student, source = 'unknown') {
    lastSelectionSourceRef.current = source;
    setActiveStudent(student);
  }

  // Prevent duplicate logs for the same student+source (React StrictMode and multiple setters can cause repeated effects)
  // renamed for clarity
  const lastLoggedSelectionRef = useRef({ key: null, source: null });

  // Log activeStudent once when it actually changes (avoids render-time and StrictMode duplicates)
  useEffect(() => {
    try {
      // compute a small stable key for deduplication
      const key = activeStudent
        ? (activeStudent.ID || activeStudent.id || activeStudent.StudentId || activeStudent.StudentID || JSON.stringify(activeStudent))
        : '<<null>>';
      const source = lastSelectionSourceRef.current || 'unknown';
      // if we've already logged this exact key+source recently, skip logging
      if (lastLoggedSelectionRef.current.key === key && lastLoggedSelectionRef.current.source === source) {
        // nothing to do
      } else {
        if (activeStudent) {
          console.log(`activeStudent (source: ${source}):`, activeStudent);
          // load and log "Student History" sheet once when a student is selected
          try {
            loadSheet('Student History','Student',activeStudent.StudentName)
              .then(res => console.log(`loadSheet('Student History','Student','${activeStudent.StudentName}')`, res))
              .catch(err => console.error(`loadSheet('Student History','Student','${activeStudent.StudentName}') error:`, err));
          } catch (_) { /* ignore */ }
        } else {
          console.log(`activeStudent cleared (source: ${source})`);
        }
        lastLoggedSelectionRef.current = { key, source };
      }
    } catch (_) {}
    lastSelectionSourceRef.current = null;
  }, [activeStudent]);

  // Initialize currentUserName from cache/SSO on mount
  useEffect(() => {
    try {
      const cached = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
      if (cached) {
        setCurrentUserName(cached);
        sessionUserRef.current = cached;
        return;
      }
      if (window.SSO && typeof window.SSO.getUserName === 'function') {
        const n = window.SSO.getUserName();
        if (n) {
          setCurrentUserName(n);
          sessionUserRef.current = n;
        }
      }
    } catch (_) { /* ignore */ }
  }, []);

  // Persist currentUserName and keep sessionUserRef in sync
  useEffect(() => {
    if (!currentUserName) return;
    try { window.localStorage.setItem('ssoUserName', currentUserName); } catch (_) {}
    sessionUserRef.current = currentUserName;
  }, [currentUserName]);

  // Error handler
  const errorHandler = (error) => {
    console.error("onWorksheetChanged error:", error);
  };

  // renamed to clarify this flag guards the worksheet-change handler
  const isWorksheetHandlerRunning = useRef(false);

  // Use a stable handler that reads sessionUserRef to avoid stale closures
  async function onWorksheetChanged(eventArgs) {
    if (typeof window.Excel === "undefined") return;
    if (!eventArgs || !eventArgs.address) return;
    if (isWorksheetHandlerRunning.current) return;
    isWorksheetHandlerRunning.current = true;
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
            const effectiveUser = sessionUserRef.current || currentUserName || window.localStorage.getItem('ssoUserName') || 'Unknown';
            // If the comment text matches a trigger, tag as 'Contacted' and highlight the row.
            if (isOutreachTrigger(newValue)) {
              // await the comment insert, then highlight
              try {
                await addComment(newValue, 'Contacted, Outreach', effectiveUser, studentId, studentName);
              } catch (_) { /* ignore insert errors */ }
              try {
                await highlightRow(
                  rowIndex,
                  Math.min(studentNameColIndex, outreachColIndex),
                  Math.abs(studentNameColIndex - outreachColIndex) + 1,
                  'yellow'
                );
              } catch (_) { /* ignore highlight errors */ }
            } else {
              // await insert so UI may update elsewhere
              try {
                await addComment(newValue, 'Outreach', effectiveUser, studentId, studentName);
              } catch (_) { /* ignore errors */ }
            }
          }
        }
      });
    } catch (error) {
      errorHandler(error);
    } finally {
      isWorksheetHandlerRunning.current = false;
    }
  }

  // Effect: Load sheet cache (Excel or test-mode) once
  useEffect(() => {
    let mounted = true;

    const applyResult = (res) => {
      if (!mounted) return;
      setSheetData({ status: res.status || 'success', data: res.data || {}, message: res.message || '' });
      setHeaders(res.headers || []);
      setAssignmentsMap(res.assignmentsMap || {});
      if (res.status === 'success' && (!window.Excel || Object.keys(res.data || {}).length === 1)) {
        const firstKey = Object.keys(res.data || [])[0];
        if (firstKey)
          // use the renamed setter here
          setSelectedStudentWithSource(res.data[firstKey], 'initialLoad');
      }
    };

    loadCache()
      .then(applyResult)
      .catch(() => {
        applyResult({ status: 'error', data: null, message: 'An error occurred while loading the data. Please try again.' });
      });

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
          // use the renamed setter here
          setSelectedStudentWithSource(sheetData.data[rowIndex] || null, 'selection');

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
  // renamed for clarity
  const runningInBrowser = typeof window.Excel === "undefined";

  // Always show SSO first until currentUserName is set
  if (!currentUserName) {
    return (
      <div className="studentview-outer">
        <SSO onNameSelect={setCurrentUserName} />
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
      className={runningInBrowser ? "studentview-outer testmode" : "studentview-outer"}
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