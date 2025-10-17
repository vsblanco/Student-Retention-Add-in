import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
import StudentHeader from './StudentHeader.jsx';
import ExampleStudent from '../utility/ExampleStudent.jsx';
import StudentAssignments from './StudentAssignments.jsx';
import { formatName } from '../utility/Conversion.jsx';
import { onSelectionChanged } from '../utility/ExcelAPI.jsx';
import SSO from '../utility/SSO.jsx';
import { COLUMN_ALIASES, COLUMN_ALIASES_ASSIGNMENTS, COLUMN_ALIASES_HISTORY, Sheets } from '../utility/ColumnMapping.jsx';
import { normalizeHeader, canonicalHeaderMap, canonicalAssignmentsHeaderMap, canonicalHistoryHeaderMap, getCanonicalName, getCanonicalColIdx } from '../utility/CanonicalMap.jsx';
import './StudentView.css';

// Extracts URL from Excel HYPERLINK formula or returns plain URL if provided
function extractHyperlink(formulaOrValue) {
  if (!formulaOrValue) return null;
  if (typeof formulaOrValue !== 'string') return null;
  const value = formulaOrValue.trim();

  // If it's a HYPERLINK formula, try to capture the first argument (the URL)
  // Examples: =HYPERLINK("https://...","Text") or =HYPERLINK('https://...','Text')
  const hyperlinkRegex = /=\s*HYPERLINK\s*\(\s*["']?([^"',)]+)["']?\s*,/i;
  const match = value.match(hyperlinkRegex);
  if (match && match[1]) {
    return match[1].trim();
  }

  // If it's a plain URL, return it
  if (/^https?:\/\//i.test(value)) {
    return value;
  }

  return null;
}

// Helper function to process a raw row of data against headers
const processRow = (rowData, headers, formulaRowData) => {
  if (!rowData || rowData.every(cell => cell === "")) {
    return null;
  }
  
  const studentInfo = {};
  let gradebookIndex = -1;

  headers.forEach((header, index) => {
    if (header && index < rowData.length) {
      // --- Normalize the header from the sheet in the same way before looking it up ---
      const normalizedLookupKey = normalizeHeader(header);
      const canonicalHeader = getCanonicalName(canonicalHeaderMap, header);
      studentInfo[canonicalHeader] = rowData[index];

      // Check if this is the GradeBook column
      if (canonicalHeader === 'Gradebook') {
          gradebookIndex = index;
      }
    }
  });

  // --- Extract GradeBook Hyperlink if available ---
  if (gradebookIndex !== -1 && formulaRowData && gradebookIndex < formulaRowData.length) {
      const formula = formulaRowData[gradebookIndex];
      const link = extractHyperlink(formula);
      if (link) {
          studentInfo.Gradebook = link;
      }
  }

  // Now we only need to check for the one, standard "StudentName" key
  if (!studentInfo.StudentName || String(studentInfo.StudentName).trim() === "") {
    return null;
  }
  return studentInfo;
};

// --- Helper: generic row processor using a provided canonical header map ---
const processRowsWithCanonicalMap = (rows, headers, canonicalMap) => {
  const headerIndexMap = {};
  headers.forEach((header, idx) => {
    const canonical = getCanonicalName(canonicalMap, header);
    if (canonical) headerIndexMap[canonical] = idx;
  });
  return rows.map(row => {
    const entry = {};
    Object.entries(headerIndexMap).forEach(([canonical, idx]) => {
      entry[canonical] = row[idx];
    });
    return entry;
  });
};

function StudentView() {
  const [activeStudent, setActiveStudent] = useState(null);
  // --- UPDATE: Enhanced state to provide better user feedback ---
  const [sheetData, setSheetData] = useState({ status: 'loading', data: null, message: 'Loading student data...' });
  const [activeTab, setActiveTab] = useState('details');
  const isInitialLoad = useRef(true);

  // --- Store student row indices for navigation ---
  const [studentRowIndices, setStudentRowIndices] = useState([]);
  const [currentRowIndex, setCurrentRowIndex] = useState(null);
  const [pendingRowIdx, setPendingRowIdx] = useState(null); // index in cache for fast navigation
  const [studentCache, setStudentCache] = useState([]); // ordered cache of students

  // --- Store headers for column lookup ---
  const [headers, setHeaders] = useState([]);
  const [assignmentsMap, setAssignmentsMap] = useState({}); // { studentName: [assignments] }
  const [userName, setUserName] = useState(null);

  // --- TEST MODE: Provide a basic student object if not running in Office/Excel ---
  useEffect(() => {
    if (typeof window.Excel === "undefined") {
      // Use imported ExampleStudent for testing
      setSheetData({ status: 'success', data: { 1: ExampleStudent }, message: '' });
      setActiveStudent(ExampleStudent);
    }
  }, []);

  // Effect 1: Load the entire sheet data into memory ONCE.
  useEffect(() => {
    if (typeof window.Excel === "undefined") return;
    const loadEntireSheet = async () => {
      try {
        await Excel.run(async (context) => {
          // --- Load student sheet ---
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const usedRange = sheet.getUsedRange(true);
          usedRange.load(["values", "formulas", "rowIndex"]);
          // --- Load history sheet by name using STUDENT_HISTORY_SHEET ---
          const historySheet = context.workbook.worksheets.getItem(Sheets.HISTORY);
          const historyRange = historySheet.getUsedRange(true);
          historyRange.load(["values"]);
          // --- Load assignments sheet ---
          let assignmentsRange, assignmentsValues = [], assignmentsHeaders = [];
          try {
            const assignmentsSheet = context.workbook.worksheets.getItem(Sheets.MISSING_ASSIGNMENT);
            assignmentsRange = assignmentsSheet.getUsedRange(true);
            assignmentsRange.load(["values"]);
          } catch (e) {
            assignmentsRange = null;
          }
          await context.sync();

          // --- Process assignments sheet ---
          let assignmentsMap = {};
          if (assignmentsRange && assignmentsRange.values && assignmentsRange.values.length > 1) {
            assignmentsHeaders = assignmentsRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
            assignmentsValues = assignmentsRange.values.slice(1);

            // Build header index map using COLUMN_ALIASES_ASSIGNMENTS
            const headerIndexMap = {};
            assignmentsHeaders.forEach((header, idx) => {
              const normalized = normalizeHeader(header);
              const canonical = getCanonicalName(canonicalAssignmentsHeaderMap, header);
              if (canonical) headerIndexMap[canonical] = idx;
            });

            assignmentsValues.forEach(row => {
              const nameIdx = headerIndexMap.StudentName;
              if (nameIdx === undefined) return;
              const studentName = row[nameIdx];
              if (!studentName) return;
              // Build assignment object using headerIndexMap
              const assignment = {
                title: row[headerIndexMap.title] || '',
                dueDate: row[headerIndexMap.dueDate] || '',
                score: row[headerIndexMap.score] || '',
                submissionLink: row[headerIndexMap.submissionLink] || '',
                assignmentLink: row[headerIndexMap.assignmentLink] || '',
                submission: row[headerIndexMap.submission] || false
              };
              if (!assignmentsMap[studentName]) assignmentsMap[studentName] = [];
              assignmentsMap[studentName].push(assignment);
            });
          }
          setAssignmentsMap(assignmentsMap);

          // --- Process history sheet ---
          let historyMap = {};
          if (historyRange.values && historyRange.values.length > 1) {
            const historyHeaders = historyRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
            const historyRows = historyRange.values.slice(1);
            // *** MODIFIED: Use COLUMN_ALIASES_HISTORY for mapping ***
            const processedHistory = processRowsWithCanonicalMap(historyRows, historyHeaders, canonicalHistoryHeaderMap);
            // Group by Student ID (case-insensitive)
            processedHistory.forEach(entry => {
              // Try to find ID using flexible mapping
              const id =
                entry["ID"] ||
                entry["Student ID"] ||
                entry["Student identifier"] ||
                "";
              if (!id) return;
              const key = String(id).toLowerCase();
              if (!historyMap[key]) historyMap[key] = [];
              historyMap[key].push(entry);
            });
          }

          // --- Process student sheet ---
          if (!usedRange.values || usedRange.values.length < 2) {
            setSheetData({ status: 'empty', data: {}, message: 'No student data found on this sheet.' });
            return;
          }
          const headers = usedRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
          setHeaders(headers); // <-- Save headers for later use
          const studentDataMap = {};
          const startRowIndex = usedRange.rowIndex;
          for (let i = 1; i < usedRange.values.length; i++) {
            // *** MODIFIED: Pass formulas for the current row ***
            const studentInfo = processRow(usedRange.values[i], headers, usedRange.formulas[i]);
            if (studentInfo) {
              // --- Attach history if Student ID matches, else set empty array ---
              const id = studentInfo.ID || "";
              const key = String(id).toLowerCase();
              if (id && historyMap[key]) {
                studentInfo.History = historyMap[key];
              } else {
                studentInfo.History = [];
              }
              const actualRowIndex = startRowIndex + i;
              studentDataMap[actualRowIndex] = studentInfo;
            }
          }
          // --- Store row indices for navigation ---
          const indices = Object.keys(studentDataMap).map(idx => Number(idx));
          setStudentRowIndices(indices);
          // --- Build ordered cache ---
          setStudentCache(indices.map(idx => studentDataMap[idx]));
          if (Object.keys(studentDataMap).length === 0) {
            setSheetData({ status: 'empty', data: {}, message: 'No student data found on this sheet.' });
          } else {
            setSheetData({ status: 'success', data: studentDataMap, message: '' });
          }
        });
      } catch (error) {
        console.error("Error loading entire sheet:", error);
        setSheetData({ status: 'error', data: null, message: 'An error occurred while loading the data. Please try again.' });
      }
    };

    loadEntireSheet();
  }, []);

  // --- Store reference to Excel selection event handler for enable/disable ---
  const selectionHandlerRef = useRef(null);
  const selectionHandlerEnabledRef = useRef(true);

  // Effect 2: Handle selection changes from Excel.
  useEffect(() => {
    if (typeof window.Excel === "undefined") return; // Skip Excel logic in test mode
    if (sheetData.status !== 'success') return;

    let eventHandlerObj = null;

    const syncUIToSheetSelection = async () => {
      if (!selectionHandlerEnabledRef.current) return; // Ignore if disabled
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load(["rowIndex", "columnIndex"]);
          await context.sync();
          const rowIndex = range.rowIndex;
          const colIndex = range.columnIndex;
          setActiveStudent(sheetData.data[rowIndex] || null);
          setCurrentRowIndex(rowIndex);

          // --- Auto navigation: Outreach -> History, Missing Assignments -> Assignments, Phone/OtherPhone -> History ---
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

    // Use the ExcelAPI utility for selection change
    onSelectionChanged(() => {
      syncUIToSheetSelection();
    }).then(obj => {
      eventHandlerObj = obj;
      selectionHandlerRef.current = obj;
      selectionHandlerEnabledRef.current = true;
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
      selectionHandlerEnabledRef.current = true;
    };
  }, [sheetData]);

  // When Excel selection changes, update currentRowIndex and clear pendingRowIdx
  useEffect(() => {
    if (activeStudent && studentCache.length) {
      const idx = studentCache.findIndex(s => s === activeStudent);
      if (idx !== -1) {
        setCurrentRowIndex(studentRowIndices[idx]);
        setPendingRowIdx(null);
      }
      // --- Log payload only when activeStudent changes ---
      //console.log('StudentHeader payload:', activeStudent);
    }
  }, [activeStudent, studentCache, studentRowIndices]);
  
  // --- Detect test mode (browser, not Excel) ---
  const isTestMode = typeof window.Excel === "undefined";

  // --- STYLES ---
  // All styles moved to CSS file

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

  // --- UPDATE: Centralized rendering logic based on sheetData.status ---
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
    // Try both converted and raw name for robustness
    const assignments =
      assignmentsMap[convertedName] ||
      assignmentsMap[studentObj.StudentName] ||
      [];
  
    return assignments;
  };

  // --- Attach Assignments to activeStudent payload ---
  const activeStudentWithAssignments = activeStudent
    ? { ...activeStudent, Assignments: getAssignmentsForStudent(activeStudent) || [] }
    : null;

  return (
    <div
      className={isTestMode ? "studentview-outer testmode" : "studentview-outer"}
      style={{ userSelect: "none" }}
    >
      {/* activeStudent now contains GradeBookLink if a hyperlink formula was found */}
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