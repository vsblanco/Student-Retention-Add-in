// Timestamp: 2025-10-02 05:29 PM | Version: 6.2.0
import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';

// --- Alias mapping for flexible column names ---
const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'Student Number'],
  Phone: ['Phone Number', 'Contact'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor']
  // You can add more aliases for other columns here
};

// --- UPDATE: Create a reverse map that is agnostic to whitespace ---
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
      // --- UPDATE: Normalize the header from the sheet in the same way before looking it up ---
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

function StudentView() {
  const [activeStudent, setActiveStudent] = useState(null);
  const [allStudents, setAllStudents] = useState(null);
  const [activeTab, setActiveTab] = useState('details');
  const isInitialLoad = useRef(true);

  // Effect 1: Load the entire sheet data into memory ONCE.
  useEffect(() => {
    const loadEntireSheet = async () => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const usedRange = sheet.getUsedRange(true);
          usedRange.load(["values", "rowIndex"]);
          await context.sync();

          if (!usedRange.values || usedRange.values.length < 2) {
            setAllStudents({}); // Set to empty object if no data
            return;
          }

          const headers = usedRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
          const studentDataMap = {};
          const startRowIndex = usedRange.rowIndex;
          
          for (let i = 1; i < usedRange.values.length; i++) {
            const studentInfo = processRow(usedRange.values[i], headers);
            if (studentInfo) {
              const actualRowIndex = startRowIndex + i;
              studentDataMap[actualRowIndex] = studentInfo; 
            }
          }
          setAllStudents(studentDataMap);
        });
      } catch (error) {
        console.error("Error loading entire sheet:", error);
        setAllStudents({});
      }
    };

    loadEntireSheet();
  }, []);

  // Effect 2: Handle selection changes from Excel.
  useEffect(() => {
    const syncUIToSheetSelection = async () => {
      if (allStudents === null) return;

      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load("rowIndex");
          await context.sync();
          
          const rowIndex = range.rowIndex;
          setActiveStudent(allStudents[rowIndex] || null);
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
    };
  }, [allStudents]);

  // --- STYLES ---
  const containerStyles = {
    padding: '0 15px 15px 15px',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
  };

  const noStudentStyles = {
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

  if (allStudents === null) {
    return <div style={noStudentStyles}><p>Loading student data...</p></div>;
  }

  if (!activeStudent) {
    return (
      <div style={noStudentStyles}>
        <p>Select a cell in a student's row to view their details.</p>
      </div>
    );
  }

  return (
    <div>
      <div style={tabContainerStyles}>
        <div style={getTabStyles('details')} onClick={() => setActiveTab('details')}>Details</div>
        <div style={getTabStyles('history')} onClick={() => setActiveTab('history')}>History</div>
      </div>
      <div style={containerStyles}>
        {activeTab === 'details' && <StudentDetails student={activeStudent} />}
        {activeTab === 'history' && <StudentHistory history={activeStudent.History} />}
      </div>
    </div>
  );
}

export default StudentView;

