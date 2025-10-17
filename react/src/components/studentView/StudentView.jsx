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

function StudentView() {
  const [activeStudent, setActiveStudent] = useState(null);
  // keep minimal UI state
  const [sheetData, setSheetData] = useState({ status: 'loading', data: null, message: 'Loading student data...' });
  const [activeTab, setActiveTab] = useState('details');
  const isInitialLoad = useRef(true);

  // keep headers and assignments map only
  const [headers, setHeaders] = useState([]);
  const [assignmentsMap, setAssignmentsMap] = useState({});
  const [userName, setUserName] = useState(null);

  // Effect: Load sheet cache (Excel or test-mode) once
  useEffect(() => {
    let mounted = true;
    const run = async () => {
      setSheetData({ status: 'loading', data: null, message: 'Loading student data...' });
      try {
        const res = await loadCache();
        if (!mounted) return;
        // res should contain: status, data, message, headers, assignmentsMap
        setSheetData({ status: res.status || 'success', data: res.data || {}, message: res.message || '' });
        setHeaders(res.headers || []);
        setAssignmentsMap(res.assignmentsMap || {});
        // Optionally set an initial active student in test-mode
        if (res.status === 'success' && (!window.Excel || Object.keys(res.data || {}).length === 1)) {
          const firstKey = Object.keys(res.data || [])[0];
          if (firstKey) setActiveStudent(res.data[firstKey]);
        }
      } catch (err) {
        console.error('Cache load error', err);
        if (!mounted) return;
        setSheetData({ status: 'error', data: null, message: 'An error occurred while loading the data. Please try again.' });
      }
    };
    run();
    return () => { mounted = false; };
  }, []);

  // Keep selection handler ref for cleanup
  const selectionHandlerRef = useRef(null);

  // Effect: Handle selection changes from Excel.
  useEffect(() => {
    if (typeof window.Excel === "undefined") return; // Skip Excel logic in test mode
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
          setActiveStudent(sheetData.data[rowIndex] || null);

          // Auto navigation: Outreach/Phone -> History, Missing Assignments -> Assignments
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

    return () => {
      if (eventHandlerObj && eventHandlerObj.remove) {
        eventHandlerObj.remove();
      }
      selectionHandlerRef.current = null;
    };
  }, [sheetData, headers]);

  // Detect test mode (browser, not Excel)
  const isTestMode = typeof window.Excel === "undefined";

  // Helper to get the correct history as an array of objects from the activeStudent object
  const getHistoryArray = (studentObj) => {
    if (!studentObj) return [];
    if (Array.isArray(studentObj.History)) return studentObj.History;
    if (typeof studentObj.History === 'string') {
      const lines = studentObj.History.split('\n').map(l => l.trim()).filter(Boolean);
      const entries = [];
      for (let i = 0; i < lines.length; i += 2) {
        entries.push({
          timestamp: lines[i] || '',
          comment: lines[i + 1] || '',
          studentId: studentObj.ID || '',
          studentName: studentObj.StudentName || '',
          tag: ''
        });
      }
      return entries;
    }
    return [];
  };

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

  // Helper to get assignments for the active student by name
  const getAssignmentsForStudent = (studentObj) => {
    if (!studentObj || !studentObj.StudentName) return [];
    const convertedName = formatName(studentObj.StudentName);
    const assignments =
      assignmentsMap[convertedName] ||
      assignmentsMap[studentObj.StudentName] ||
      [];
    return assignments;
  };

  // Attach Assignments to activeStudent payload
  const activeStudentWithAssignments = activeStudent
    ? { ...activeStudent, Assignments: getAssignmentsForStudent(activeStudent) || [] }
    : null;

  return (
    <div
      className={isTestMode ? "studentview-outer testmode" : "studentview-outer"}
      style={{ userSelect: "none" }}
    >
      <StudentHeader student={activeStudentWithAssignments} />
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
        {activeTab === 'details' && <StudentDetails student={activeStudentWithAssignments} />}
        {activeTab === 'history' && (
          <StudentHistory
            history={getHistoryArray(activeStudentWithAssignments)}
            student={activeStudentWithAssignments}
          />
        )}
        {activeTab === 'assignments' && (
          <StudentAssignments student={activeStudentWithAssignments} />
        )}
      </div>
    </div>
  );
}

export default StudentView;