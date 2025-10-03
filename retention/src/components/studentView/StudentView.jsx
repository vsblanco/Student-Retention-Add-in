// Timestamp: 2025-10-03 11:03 AM | Version: 6.3.0
import React, { useState, useEffect, useRef } from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
import StudentHeader from './StudentHeader.jsx';

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

function StudentView() {
  const [activeStudent, setActiveStudent] = useState(null);
  // --- UPDATE: Enhanced state to provide better user feedback ---
  const [sheetData, setSheetData] = useState({ status: 'loading', data: null, message: 'Loading student data...' });
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
            setSheetData({ status: 'empty', data: {}, message: 'No student data found on this sheet.' });
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

  // --- STYLES ---
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

  // --- UPDATE: Centralized rendering logic based on sheetData.status ---
  if (sheetData.status !== 'success') {
    return <div style={messageContainerStyles}><p>{sheetData.message}</p></div>;
  }

  if (!activeStudent) {
    return (
      <div style={messageContainerStyles}>
        <p>Select a cell in a student's row to view their details.</p>
      </div>
    );
  }

  return (
    <div className='Header'>
      <StudentHeader /> 
    </div>
  );
}

export default StudentView;