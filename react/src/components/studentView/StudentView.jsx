// Timestamp: 2025-10-03 11:03 AM | Version: 6.3.0
import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
import StudentHeader from './StudentHeader.jsx';

// --- Constant for Student History sheet name ---
const STUDENT_HISTORY_SHEET = "Student History";

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
  DaysOut: ['Days Out']
  // You can add more aliases for other columns here
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

// Helper function to process a raw row of data against headers
const processRow = (rowData, headers) => {
  if (!rowData || rowData.every(cell => cell === "")) {
    return null;
  }
  const studentInfo = {};
  headers.forEach((header, index) => {
    if (header && index < rowData.length) {
      // --- Normalize the header from the sheet in the same way before looking it up ---
      const normalizedLookupKey = normalizeHeader(header);
      const canonicalHeader = canonicalHeaderMap[normalizedLookupKey] || header.trim();
      studentInfo[canonicalHeader] = rowData[index];
    }
  });

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

  // --- TEST MODE: Provide a basic student object if not running in Office/Excel ---
  useEffect(() => {
    if (typeof window.Excel === "undefined") {
      // Simulate a single student row for testing
      const testStudent = {
        StudentName: "Jane Doe",
        ID: "123456",
        Gender: "Boy",
        Phone: "555-1234",
        OtherPhone: "555-5678",
        StudentEmail: "jane.doe@university.edu",
        PersonalEmail: "jane.doe@gmail.com",
        Assigned: "Dr. Smith",
        DaysOut: 10,
        Grade: "77%",
        LDA: "2024-05-20",
        // History as an array of objects for test mode
        History: [
          {
            timestamp: "2024-01-15",
            comment: "Advised: Discussed course selection.",
            ID: "123456",
            studentName: "Jane Doe",
            createdBy: "Dr. Smith",
            tag: "Outreach"
          },
          {
            timestamp: "2024-01-15",
            comment: "Left Voicemail",
            ID: "123456",
            studentName: "Jane Doe",
            createdBy: "Dr. Smith",
            
          },
          {
            timestamp: "2024-03-10",
            comment: "Follow-up: Checked on progress.",
            ID: "123456",
            studentName: "Jane Doe",
            createdBy: "Dr. Smith",
            tag: "Contacted"
          }
        ]
        // Optionally add an alias for testing alias logic:
        // Notes: [...]
      };
      setSheetData({ status: 'success', data: { 1: testStudent }, message: '' });
      setActiveStudent(testStudent);
    }
  }, []);

  // Effect 1: Load the entire sheet data into memory ONCE.
  useEffect(() => {
    if (typeof window.Excel === "undefined") return; // Skip Excel logic in test mode
    const loadEntireSheet = async () => {
      try {
        await Excel.run(async (context) => {
          // --- Load student sheet ---
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const usedRange = sheet.getUsedRange(true);
          usedRange.load(["values", "rowIndex"]);
          // --- Load history sheet by name using STUDENT_HISTORY_SHEET ---
          const historySheet = context.workbook.worksheets.getItem(STUDENT_HISTORY_SHEET);
          const historyRange = historySheet.getUsedRange(true);
          historyRange.load(["values"]);
          await context.sync();

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
          const studentDataMap = {};
          const startRowIndex = usedRange.rowIndex;
          for (let i = 1; i < usedRange.values.length; i++) {
            const studentInfo = processRow(usedRange.values[i], headers);
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

  // Effect 2: Handle selection changes from Excel.
  useEffect(() => {
    if (typeof window.Excel === "undefined") return; // Skip Excel logic in test mode

    // Do not proceed if data isn't loaded successfully
    if (sheetData.status !== 'success') return;

    const syncUIToSheetSelection = async () => {
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load("rowIndex");
          await context.sync();
          
          const rowIndex = range.rowIndex;
          setActiveStudent(sheetData.data[rowIndex] || null);
        });
      } catch (error) {
        // Ignore GeneralException which can occur if selection is invalid
        if (error.code !== "GeneralException") {
             console.error("Error in syncUIToSheetSelection:", error);
        }
      }
    };

    let eventHandler;
    Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        eventHandler = worksheet.onSelectionChanged.add(syncUIToSheetSelection);
        await context.sync();
        
        // On initial load, manually trigger sync to reflect current selection
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
    };
  }, [sheetData]); // Depend on the entire sheetData object

  // Log payload only when activeStudent changes
  useEffect(() => {
    if (activeStudent) {
      console.log('StudentHeader payload:', activeStudent);
    }
  }, [activeStudent]);

  // --- Detect test mode (browser, not Excel) ---
  const isTestMode = typeof window.Excel === "undefined";

  // --- STYLES ---
  const outerContainerStyles = isTestMode
    ? {
        maxWidth: '400px',
        minWidth: '320px',
        margin: '0 auto',
        border: '1px solid #e5e7eb',
        borderRadius: '12px',
        boxShadow: '0 2px 12px rgba(0,0,0,0.06)',
        background: '#fff',
        minHeight: '100vh',
        overflow: 'hidden'
      }
    : {};

  const containerStyles = {
    padding: '0 15px 15px 15px',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
  };

  const messageContainerStyles = {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '80vh',
    color: '#7f8c8d',
    textAlign: 'center',
    padding: '15px'
  };

  const tabContainerStyles = {
    display: 'flex',
    borderBottom: '2px solid #e1e1e1',
    marginBottom: '15px'
  };

  const getTabStyles = (tabName) => ({
    padding: '10px 15px',
    cursor: 'pointer',
    borderBottom: activeTab === tabName ? '2px solid #2c3e50' : '2px solid transparent',
    color: activeTab === tabName ? '#2c3e50' : '#7f8c8d',
    fontWeight: activeTab === tabName ? '600' : '400',
    marginRight: '10px',
    transition: 'all 0.2s ease-in-out'
  });

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
      <div style={outerContainerStyles}>
        <div style={messageContainerStyles}><p>{sheetData.message}</p></div>
      </div>
    );
  }

  if (!activeStudent) {
    return (
      <div style={outerContainerStyles}>
        <div style={messageContainerStyles}>
          <p>Select a cell in a student's row to view their details.</p>
        </div>
      </div>
    );
  }

  return (
    <div style={outerContainerStyles}>
      <StudentHeader student={activeStudent} /> 
      <div style={tabContainerStyles}>
        <div style={getTabStyles('details')} onClick={() => setActiveTab('details')}>Details</div>
        <div style={getTabStyles('history')} onClick={() => setActiveTab('history')}>History</div>
      </div>
      <div style={containerStyles}>
        {activeTab === 'details' && <StudentDetails student={activeStudent} />}
        {activeTab === 'history' && (
          <StudentHistory history={getHistoryArray(activeStudent)} />
        )}
      </div>
    </div>
  );
}

export default StudentView;