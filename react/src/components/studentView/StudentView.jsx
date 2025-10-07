import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
import StudentHeader from './StudentHeader.jsx';
import ExampleStudent from '../utility/ExampleStudent.jsx';
import StudentAssignments from './StudentAssignments.jsx';
import { formatName } from '../utility/Conversion.jsx';
import './StudentView.css';
import ContextMenu from '../utility/ContextMenu.jsx';

const STUDENT_HISTORY_SHEET = "Student History";
const STUDENT_MISSING_ASSIGNMENT_SHEET = "Missing Assignments";

// --- Alias mapping for flexible column names ---
const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'Student Number','Student identifier'],
  Gender: ['Gender'],
  Phone: ['Phone Number', 'Contact'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor'],
  Grade: ['Current Grade', 'Grade %', 'Grade'],
  LDA: ['Last Date of Attendance', 'LDA'],
  DaysOut: ['Days Out'],
  Gradebook: ['Gradebook']
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

  // --- NEW LOGIC: Extract GradeBook Hyperlink if available ---
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

// --- Helper to process history rows ---
const processHistoryRows = (rows, headers) => {
  return rows.map(row => {
    const entry = {};
    headers.forEach((header, idx) => {
      entry[header] = row[idx];
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
            // Debug: log assignments map after processing
            console.log('[DEBUG] assignmentsMap:', assignmentsMap);
          }
          setAssignmentsMap(assignmentsMap);

          // --- Process history sheet ---
          let historyMap = {};
          if (historyRange.values && historyRange.values.length > 1) {
            const historyHeaders = historyRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
            const historyRows = historyRange.values.slice(1);
            const processedHistory = processHistoryRows(historyRows, historyHeaders);
            // Group by Student ID (case-insensitive)
            processedHistory.forEach(entry => {
              const id = entry["Student ID"] || entry["ID"] || entry["Student identifier"] || "";
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

    const syncUIToSheetSelection = async () => {
      if (!selectionHandlerEnabledRef.current) return; // Ignore if disabled
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load("rowIndex");
          await context.sync();
          const rowIndex = range.rowIndex;
          setActiveStudent(sheetData.data[rowIndex] || null);
          setCurrentRowIndex(rowIndex);
        });
      } catch (error) {
        if (error.code !== "GeneralException") {
          console.error("Error in syncUIToSheetSelection:", error);
        }
      }
    };

    let eventHandler;
    Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      eventHandler = worksheet.onSelectionChanged.add(syncUIToSheetSelection);
      selectionHandlerRef.current = eventHandler;
      selectionHandlerEnabledRef.current = true;
      await context.sync();
      if (isInitialLoad.current) {
        syncUIToSheetSelection();
        isInitialLoad.current = false;
      }
    });

    return () => {
      if (eventHandler) {
        Excel.run(async (context) => {
          eventHandler.remove();
          await context.sync();
        });
      }
      selectionHandlerRef.current = null;
      selectionHandlerEnabledRef.current = true;
    };
  }, [sheetData]);

  // --- Keyboard navigation for up/down arrows (hold support) ---
  useEffect(() => {
    if (typeof window.Excel === "undefined") return; // Only in Excel

    let navTimer = null;
    let lastDirection = null;
    let isNavigating = false; // Prevent overlapping Excel.run calls

    const NAV_INTERVAL = 250; // ms between moves when holding (increased for Excel)

    const moveSelection = async (direction) => {
      if (isNavigating) return;
      isNavigating = true;
      if (!studentRowIndices.length || currentRowIndex === null) {
        isNavigating = false;
        return;
      }
      const currentIdx = studentRowIndices.indexOf(currentRowIndex);
      let nextIdx = currentIdx;
      if (direction === "up") {
        nextIdx = Math.max(0, currentIdx - 1);
      } else if (direction === "down") {
        nextIdx = Math.min(studentRowIndices.length - 1, currentIdx + 1);
      }
      const nextRowIndex = studentRowIndices[nextIdx];
      if (nextRowIndex !== undefined && nextRowIndex !== currentRowIndex) {
        selectionHandlerEnabledRef.current = false; // Disable selection event handler
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const nameColIdx = getStudentNameColIdx();
          const range = sheet.getRangeByIndexes(nextRowIndex, nameColIdx, 1, 1);
          range.select();
          await context.sync();
          setActiveStudent(sheetData.data[nextRowIndex] || null);
          setCurrentRowIndex(nextRowIndex);
        });
        selectionHandlerEnabledRef.current = true; // Re-enable after navigation
      }
      isNavigating = false;
    };

    const handleKeyDown = (e) => {
      if (e.repeat && navTimer) return; // Already repeating
      if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        e.preventDefault();
        lastDirection = e.key === "ArrowUp" ? "up" : "down";
        moveSelection(lastDirection);
        // Start timer for hold
        if (!navTimer) {
          navTimer = setInterval(() => moveSelection(lastDirection), NAV_INTERVAL);
        }
      }
    };

    const handleKeyUp = (e) => {
      if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        if (navTimer) {
          clearInterval(navTimer);
          navTimer = null;
        }
        lastDirection = null;
      }
    };

    const handleBlur = () => {
      if (navTimer) {
        clearInterval(navTimer);
        navTimer = null;
      }
      lastDirection = null;
    };

    window.addEventListener("keydown", handleKeyDown);
    window.addEventListener("keyup", handleKeyUp);
    window.addEventListener("blur", handleBlur);

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      window.removeEventListener("keyup", handleKeyUp);
      window.removeEventListener("blur", handleBlur);
      if (navTimer) clearInterval(navTimer);
    };
  }, [studentRowIndices, currentRowIndex, sheetData, headers]);

  // --- Fast local navigation using cache, sync Excel selection after ---
  useEffect(() => {
    if (typeof window.Excel === "undefined") return; // Only in Excel

    let navTimer = null;
    let lastDirection = null;

    const NAV_INTERVAL = 80; // Fast UI navigation

    const moveLocalSelection = (direction) => {
      if (!studentCache.length) return;
      const idx = pendingRowIdx ?? studentRowIndices.indexOf(currentRowIndex);
      let nextIdx = idx;
      if (direction === "up") {
        nextIdx = Math.max(0, idx - 1);
      } else if (direction === "down") {
        nextIdx = Math.min(studentCache.length - 1, idx + 1);
      }
      if (nextIdx !== idx) {
        setPendingRowIdx(nextIdx);
        setActiveStudent(studentCache[nextIdx]);
      }
    };

    const handleKeyDown = (e) => {
      if (e.repeat && navTimer) return;
      if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        e.preventDefault();
        lastDirection = e.key === "ArrowUp" ? "up" : "down";
        moveLocalSelection(lastDirection);
        if (!navTimer) {
          navTimer = setInterval(() => moveLocalSelection(lastDirection), NAV_INTERVAL);
        }
      }
    };

    const handleKeyUp = (e) => {
      if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        if (navTimer) {
          clearInterval(navTimer);
          navTimer = null;
        }
        lastDirection = null;
        // Sync Excel selection to UI
        const idx = pendingRowIdx ?? studentRowIndices.indexOf(currentRowIndex);
        if (idx !== null && idx !== undefined && idx !== studentRowIndices.indexOf(currentRowIndex)) {
          const rowIdx = studentRowIndices[idx];
          selectionHandlerEnabledRef.current = false; // Disable selection event handler
          Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const nameColIdx = getStudentNameColIdx();
            const range = sheet.getRangeByIndexes(rowIdx, nameColIdx, 1, 1);
            range.select();
            await context.sync();
            setCurrentRowIndex(rowIdx);
          }).finally(() => {
            selectionHandlerEnabledRef.current = true; // Re-enable after navigation
          });
        }
        setPendingRowIdx(null);
      }
    };

    const handleBlur = () => {
      if (navTimer) {
        clearInterval(navTimer);
        navTimer = null;
      }
      lastDirection = null;
      setPendingRowIdx(null);
    };

    window.addEventListener("keydown", handleKeyDown);
    window.addEventListener("keyup", handleKeyUp);
    window.addEventListener("blur", handleBlur);

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      window.removeEventListener("keyup", handleKeyUp);
      window.removeEventListener("blur", handleBlur);
      if (navTimer) clearInterval(navTimer);
    };
  }, [studentCache, studentRowIndices, currentRowIndex, pendingRowIdx, headers]);

  // When Excel selection changes, update currentRowIndex and clear pendingRowIdx
  useEffect(() => {
    if (activeStudent && studentCache.length) {
      const idx = studentCache.findIndex(s => s === activeStudent);
      if (idx !== -1) {
        setCurrentRowIndex(studentRowIndices[idx]);
        setPendingRowIdx(null);
      }
      // --- Log payload only when activeStudent changes ---
      // This log now includes the GradeBookLink property!
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
          createdBy: studentObj.Assigned || '',
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
    // Debug: log assignments for current student
    console.log('[DEBUG] Assignments for', studentObj.StudentName, assignments);
    return assignments;
  };

  // --- Attach Assignments to activeStudent payload ---
  const activeStudentWithAssignments = activeStudent
    ? { ...activeStudent, Assignments: getAssignmentsForStudent(activeStudent) || [] }
    : null;

  // Debug: log payload sent to StudentAssignments
  if (activeTab === 'assignments' && activeStudentWithAssignments) {
    console.log('[DEBUG] StudentAssignments payload:', activeStudentWithAssignments);
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
