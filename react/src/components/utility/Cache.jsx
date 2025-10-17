import ExampleStudent from './ExampleStudent.jsx';
import { normalizeHeader, getCanonicalName, canonicalHeaderMap } from './CanonicalMap.jsx'; // <-- added canonicalHeaderMap
import { canonicalAssignmentsHeaderMap, canonicalHistoryHeaderMap } from './CanonicalMap.jsx';
import { Sheets } from './ColumnMapping.jsx';

// Helper: extract Hyperlink from Excel HYPERLINK formula or plain URL
function extractHyperlink(formulaOrValue) {
  if (!formulaOrValue) return null;
  if (typeof formulaOrValue !== 'string') return null;
  const value = formulaOrValue.trim();
  const hyperlinkRegex = /=\s*HYPERLINK\s*\(\s*["']?([^"',)]+)["']?\s*,/i;
  const match = value.match(hyperlinkRegex);
  if (match && match[1]) {
    return match[1].trim();
  }
  if (/^https?:\/\//i.test(value)) {
    return value;
  }
  return null;
}

// Helper function to process a raw row of data against headers (same logic moved from StudentView)
const processRow = (rowData, headers, formulaRowData) => {
  if (!rowData || rowData.every(cell => cell === "")) {
    return null;
  }
  
  const studentInfo = {};
  let gradebookIndex = -1;

  headers.forEach((header, index) => {
    if (header && index < rowData.length) {
      const canonicalHeader = getCanonicalName(canonicalHeaderMap, header) || header; // use canonicalHeaderMap
      studentInfo[canonicalHeader] = rowData[index];

      if (canonicalHeader === 'Gradebook') {
          gradebookIndex = index;
      }
    }
  });

  if (gradebookIndex !== -1 && formulaRowData && gradebookIndex < formulaRowData.length) {
      const formula = formulaRowData[gradebookIndex];
      const link = extractHyperlink(formula);
      if (link) {
          studentInfo.Gradebook = link;
      }
  }

  if (!studentInfo.StudentName || String(studentInfo.StudentName).trim() === "") {
    return null;
  }
  return studentInfo;
};

// Generic rows processor using a provided canonical header map
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

// Public: loadCache - loads sheets (Excel) or returns test-mode ExampleStudent
export async function loadCache() {
  // Test / browser mode
  if (typeof window.Excel === "undefined") {
    return {
      status: 'success',
      message: '',
      data: { 1: ExampleStudent },
      studentCache: [ExampleStudent],
      studentRowIndices: [1],
      headers: [],
      assignmentsMap: {}
    };
  }

  try {
    const result = await Excel.run(async (context) => {
      // --- Load student sheet ---
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange(true);
      usedRange.load(["values", "formulas", "rowIndex"]);

      // --- Load history sheet ---
      const historySheet = context.workbook.worksheets.getItem(Sheets.HISTORY);
      const historyRange = historySheet.getUsedRange(true);
      historyRange.load(["values"]);

      // --- Load assignments sheet (optional) ---
      let assignmentsRange = null;
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
        const assignmentsHeaders = assignmentsRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
        const assignmentsValues = assignmentsRange.values.slice(1);

        const headerIndexMap = {};
        assignmentsHeaders.forEach((header, idx) => {
          const canonical = getCanonicalName(canonicalAssignmentsHeaderMap, header);
          if (canonical) headerIndexMap[canonical] = idx;
        });

        assignmentsValues.forEach(row => {
          const nameIdx = headerIndexMap.StudentName;
          if (nameIdx === undefined) return;
          const studentName = row[nameIdx];
          if (!studentName) return;
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

      // --- Process history sheet ---
      let historyMap = {};
      if (historyRange.values && historyRange.values.length > 1) {
        const historyHeaders = historyRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
        const historyRows = historyRange.values.slice(1);
        const processedHistory = processRowsWithCanonicalMap(historyRows, historyHeaders, canonicalHistoryHeaderMap);
        processedHistory.forEach(entry => {
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
        return {
          status: 'empty',
          message: 'No student data found on this sheet.',
          data: {},
          studentCache: [],
          studentRowIndices: [],
          headers: [],
          assignmentsMap
        };
      }

      const headers = usedRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
      const studentDataMap = {};
      const startRowIndex = usedRange.rowIndex;
      for (let i = 1; i < usedRange.values.length; i++) {
        const studentInfo = processRow(usedRange.values[i], headers, usedRange.formulas[i]);
        if (studentInfo) {
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

      const indices = Object.keys(studentDataMap).map(idx => Number(idx));
      const studentCache = indices.map(idx => studentDataMap[idx]);

      if (Object.keys(studentDataMap).length === 0) {
        return {
          status: 'empty',
          message: 'No student data found on this sheet.',
          data: {},
          studentCache: [],
          studentRowIndices: [],
          headers,
          assignmentsMap
        };
      }

      return {
        status: 'success',
        message: '',
        data: studentDataMap,
        studentCache,
        studentRowIndices: indices,
        headers,
        assignmentsMap
      };
    });

    return result;
  } catch (error) {
    console.error('loadCache error', error);
    return { status: 'error', message: 'An error occurred while loading the data. Please try again.', data: null, studentCache: [], studentRowIndices: [], headers: [], assignmentsMap: {} };
  }
}
