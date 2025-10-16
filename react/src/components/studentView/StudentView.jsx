import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
import StudentHeader from './StudentHeader.jsx';
import ExampleStudent from '../utility/ExampleStudent.jsx';
import StudentAssignments from './StudentAssignments.jsx';
import { formatName } from '../utility/Conversion.jsx';
import { onSelectionChanged } from '../utility/ExcelAPI.jsx'; // <-- Import here
import SSO from '../utility/SSO.jsx';
import './StudentView.css';

const STUDENT_HISTORY_SHEET = "Student History";
const STUDENT_MISSING_ASSIGNMENT_SHEET = "Missing Assignments";

// --- Alias mapping for flexible column names ---
const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'Student Number','Student identifier'],
  Gender: ['Gender'],
  Phone: ['Phone Number', 'Contact'],
  OtherPhone: ['Other Phone', 'Alt Phone'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor'],
  Grade: ['Current Grade', 'Grade %', 'Grade'],
  LDA: ['Last Date of Attendance', 'LDA'],
  DaysOut: ['Days Out'],
  Gradebook: ['Gradebook'],
  MissingAssignments: ['Missing Assignments', 'Missing'],
  Outreach: ['Outreach', 'Comments', 'Notes', 'Comment']
  // You can add more aliases for other columns here
};

// --- Alias mapping for flexible column names in Missing Assignments sheet ---
const COLUMN_ALIASES_ASSIGNMENTS = {
  StudentName: ['Student Name', 'Student'],
  title: ['Assignment Title', 'Title', 'Assignment'],
  dueDate: ['Due Date', 'Deadline'],
  score: ['Score', 'Points'],
  submissionLink: ['Submission Link', 'Submission', 'Submit Link'],
  assignmentLink: ['Assignment Link', 'Assignment URL', 'Assignment Page', 'Link']
};

// --- Alias mapping for flexible column names in Student History sheet ---
const COLUMN_ALIASES_HISTORY = {
  timestamp: ['Timestamp', 'Date', 'Time', 'Created At'],
  comment: ['Comment', 'Notes', 'History', 'Entry'],
  createdBy: ['Created By', 'Author', 'Advisor'],
  tag: ['Tag', 'Category', 'Type','Tags']
  // Add more aliases as needed for history columns
};

// --- Create a reverse map that is agnostic to whitespace ---
const canonicalHeaderMap = {};
// A helper function to normalize strings by removing spaces and making them lowercase
const normalizeHeader = (str) => str.toLowerCase().replace(/\s/g, '');

for (const canonicalName in COLUMN_ALIASES) {
  // Add the nicknames
  COLUMN_ALIASES[canonicalName].forEach(alias => {
    canonicalHeaderMap[normalizeHeader(alias)] = canonicalName;
  });
  // Add the canonical name itself for direct matches
  canonicalHeaderMap[normalizeHeader(canonicalName)] = canonicalName;
}

// --- Create a reverse map for assignments sheet ---
const canonicalAssignmentsHeaderMap = {};
for (const canonicalName in COLUMN_ALIASES_ASSIGNMENTS) {
  COLUMN_ALIASES_ASSIGNMENTS[canonicalName].forEach(alias => {
    canonicalAssignmentsHeaderMap[normalizeHeader(alias)] = canonicalName;
  });
  canonicalAssignmentsHeaderMap[normalizeHeader(canonicalName)] = canonicalName;
}

/**
 * Extracts the URL from an Excel HYPERLINK formula string.
 * e.g., '=HYPERLINK("https://link.com/gradebook", "Gradebook")' -> 'https://link.com/gradebook'
 * @param {string} formula - The raw Excel formula string.
 * @returns {string | null} The extracted URL or null.
 */
const extractHyperlink = (formula) => {
  if (typeof formula !== 'string' || !formula.toUpperCase().startsWith('=HYPERLINK(')) {
    return null;
  }
  const regex = /=HYPERLINK\s*\(\s*["']?([^,"']+)["']?\s*,\s*[^)]+\)/i;
  const match = formula.match(regex);
  
  if (match && match[1]) {
    let url = match[1].trim();
    // In case the captured group still has quotes (depending on Excel's formula output)
    if (url.startsWith('"') && url.endsWith('"')) {
        url = url.substring(1, url.length - 1);
    }
    if (url.startsWith("'") && url.endsWith("'")) {
        url = url.substring(1, url.length - 1);
    }
    return url;
  }
  
  return null;
};

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
      const canonicalHeader = canonicalHeaderMap[normalizedLookupKey] || header.trim();
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

// --- Helper to process history rows using COLUMN_ALIASES_HISTORY ---
const processHistoryRows = (rows, headers) => {
  // Build header index map using COLUMN_ALIASES_HISTORY
  const headerIndexMap = {};
  headers.forEach((header, idx) => {
    const normalized = normalizeHeader(header);
    const canonical = COLUMN_ALIASES_HISTORY[normalized] ? normalized : Object.keys(COLUMN_ALIASES_HISTORY).find(key =>
      COLUMN_ALIASES_HISTORY[key].map(a => normalizeHeader(a)).includes(normalized)
    );
    if (canonical) {
      // Use canonical name for mapping
      const canonicalName = canonicalAssignmentsHeaderMap[normalized] || canonical;
      headerIndexMap[canonicalName] = idx;
    }
  });

  return rows.map(row => {
    const entry = {};
    // Map each canonical history field
    Object.keys(COLUMN_ALIASES_HISTORY).forEach(canonical => {
      const idx = headerIndexMap[canonical];
      entry[canonical] = idx !== undefined ? row[idx] : '';
    });
    // Add any unmapped headers as fallback
    headers.forEach((header, idx) => {
      if (!Object.values(headerIndexMap).includes(idx)) {
        entry[header] = row[idx];
      }
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

  // --- Helper to find the column index for StudentName based on stored headers ---
  const getStudentNameColIdx = () => {
    const studentNameCanonical = 'StudentName';
    for (let i = 0; i < headers.length; i++) {
        const normalizedLookupKey = normalizeHeader(headers[i]);
        const canonicalHeader = canonicalHeaderMap[normalizedLookupKey] || headers[i].trim();
        if (canonicalHeader === studentNameCanonical) {
            return i;
        }
    }
    // Default to the first column if not found (a guess)
    return 0; 
  };

  // --- General helper to find the column index by canonical name or alias ---
  const getColIdx = (colName) => {
    // Get all possible aliases for the column, including the canonical name
    const aliases = COLUMN_ALIASES[colName] ? [colName, ...COLUMN_ALIASES[colName]] : [colName];
    const normalizedAliases = aliases.map(a => normalizeHeader(a));
    for (let i = 0; i < headers.length; i++) {
      const normalized = normalizeHeader(headers[i]);
      if (normalizedAliases.includes(normalized)) {
        return i;
      }
    }
    return -1;
  };

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
          const historySheet = context.workbook.worksheets.getItem(STUDENT_HISTORY_SHEET);
          const historyRange = historySheet.getUsedRange(true);
          historyRange.load(["values"]);
          // --- Load assignments sheet ---
          let assignmentsRange, assignmentsValues = [], assignmentsHeaders = [];
          try {
            const assignmentsSheet = context.workbook.worksheets.getItem(STUDENT_MISSING_ASSIGNMENT_SHEET);
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
              const canonical = canonicalAssignmentsHeaderMap[normalized];
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
            const processedHistory = processHistoryRows(historyRows, historyHeaders);
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
          const outreachIdx = getColIdx('Outreach');
          const missingAssignmentsIdx = getColIdx('Missing Assignments');
          const phoneIdx = getColIdx('Phone');
          const otherPhoneIdx = getColIdx('OtherPhone');
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
      console.log('StudentHeader payload:', activeStudent);
    }
  }, [activeStudent, studentCache, studentRowIndices]);
  
  // --- Detect test mode (browser, not Excel) ---
  const isTestMode = typeof window.Excel === "undefined";

  // --- STYLES ---
  // All styles moved to CSS file

  // Helper to get the correct history as an array of objects from the activeStudent object
  const getHistoryArray = (studentObj) => {
    if (!studentObj) return [];
    // If already an array, return as is
    if (Array.isArray(studentObj.History)) return studentObj.History;
    // If it's a string, parse it into objects (legacy/test mode)
    if (typeof studentObj.History === 'string') {
      // Example format: alternating lines of date and comment
      const lines = studentObj.History.split('\n').map(l => l.trim()).filter(l => l !== '');
      const entries = [];
      for (let i = 0; i < lines.length; i += 2) {
        entries.push({
          timestamp: lines[i] || '',
          comment: lines[i + 1] || '',
          // Fill with empty/defaults for other fields
          studentId: studentObj.ID || '',
          studentName: studentObj.StudentName || '',
          tag: '', // No tag in legacy/test mode
        });
      }
      return entries;
    }
    // If missing or unknown format, return empty array
    return [];
  };

  // --- UPDATE: Centralized rendering logic based on sheetData.status ---
  if (sheetData.status !== 'success') {
    return (
      <div className={isTestMode ? "studentview-outer testmode" : "studentview-outer"}>
        <div className="studentview-message"><p>{sheetData.message}</p></div>
      </div>
    );
  }

  if (!activeStudent) {
    return (
      <div className={isTestMode ? "studentview-outer testmode" : "studentview-outer"}>
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

  // Wait for userName before initializing StudentView
  if (!userName) {
    return (
      <div className="studentview-outer">
        <SSO onNameSelect={setUserName} />
      </div>
    );
  }

  return (
    <div className={isTestMode ? "studentview-outer testmode" : "studentview-outer"}>
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
          <StudentHistory history={getHistoryArray(activeStudentWithAssignments)} />
        )}
        {activeTab === 'assignments' && (
          <StudentAssignments student={activeStudentWithAssignments} />
        )}
      </div>
    </div>
  );
}

export default StudentView;
