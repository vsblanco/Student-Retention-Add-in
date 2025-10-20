import ExampleStudent from './ExampleStudent.jsx';
import { normalizeHeader, getCanonicalName, canonicalHeaderMap } from './CanonicalMap.jsx'; // <-- added canonicalHeaderMap
import { Sheets, COLUMN_ALIASES_HISTORY, COLUMN_ALIASES_ASSIGNMENTS } from './ColumnMapping.jsx';

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


// add: resilient logger to increase visibility across environments
function safeLog(...args) {
	// try normal console
	try {
		if (typeof console !== 'undefined' && typeof console.log === 'function') {
			console.log(...args);
		}
	} catch (e) { /* swallow */ }

	// also try posting to parent frame (useful if running inside a webview or iframe)
	try {
		if (typeof window !== 'undefined' && window.parent && window !== window.parent && typeof window.parent.postMessage === 'function') {
			try {
				window.parent.postMessage({ source: 'StudentRetentionAddin', log: args }, '*');
			} catch (e) { /* swallow */ }
		}
	} catch (_) { /* swallow */ }
}

// New: loadHistory - loads and processes only the HISTORY sheet and returns a map keyed by student id (lowercased)
export async function loadHistory() {
  // Test / browser mode
  if (typeof window.Excel === "undefined") {
    return {
      status: 'success',
      message: '',
      historyMap: {},
      headers: []
    };
  }

  try {
    const result = await Excel.run(async (context) => {
      try {
        const historySheet = context.workbook.worksheets.getItem(Sheets.HISTORY);
        const historyRange = historySheet.getUsedRange(true);
        historyRange.load(["values"]);
        await context.sync();

        if (!historyRange.values || historyRange.values.length < 2) {
          return {
            status: 'empty',
            message: 'No history data found on this sheet.',
            historyMap: {},
            headers: []
          };
        }

        const historyHeaders = historyRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
        const historyRows = historyRange.values.slice(1);

        // Build headerIndexMap based on COLUMN_ALIASES_HISTORY
        const headerIndexMap = {};
        historyHeaders.forEach((header, idx) => {
          if (!header) return;
          const hLower = String(header).trim().toLowerCase();
          Object.entries(COLUMN_ALIASES_HISTORY).forEach(([canonical, aliases]) => {
            const aliasList = Array.isArray(aliases) ? aliases : [aliases];
            const matches = aliasList.concat([canonical]).some(a => a && String(a).trim().toLowerCase() === hLower);
            if (matches) headerIndexMap[canonical] = idx;
          });
        });

        // Build processed history rows using headerIndexMap so only desired columns are included
        const historyMap = {};
        historyRows.forEach(row => {
          const entry = {};
          Object.entries(headerIndexMap).forEach(([canonical, idx]) => {
            entry[canonical] = row[idx];
          });

          const id =
            entry.StudentID ||
            entry.ID ||
            entry['Student ID'] ||
            entry['Student identifier'] ||
            "";
          if (!id) return;
          const key = String(id).toLowerCase();
          if (!historyMap[key]) historyMap[key] = [];
          historyMap[key].push(entry);
        });

        try { safeLog('Loaded historyMap keys:', Object.keys(historyMap).length); } catch (_) {}

        return {
          status: 'success',
          message: '',
          historyMap,
          headers: historyHeaders
        };
      } catch (e) {
        // worksheet not found or other issue
        return {
          status: 'missing',
          message: 'History sheet not found or could not be loaded.',
          historyMap: {},
          headers: []
        };
      }
    });

    return result;
  } catch (error) {
    console.error('loadHistory error', error);
    return { status: 'error', message: 'An error occurred while loading the history sheet.', historyMap: {}, headers: [] };
  }
}

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

      // --- Process assignments sheet (use COLUMN_ALIASES_ASSIGNMENTS for column selection) ---
      let assignmentsMap = {};
      if (assignmentsRange && assignmentsRange.values && assignmentsRange.values.length > 1) {
        const assignmentsHeaders = assignmentsRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
        const assignmentsRows = assignmentsRange.values.slice(1);

        // Build headerIndexMap: canonicalKey -> column index (use COLUMN_ALIASES_ASSIGNMENTS)
        const headerIndexMap = {};
        assignmentsHeaders.forEach((header, idx) => {
          if (!header) return;
          const hLower = String(header).trim().toLowerCase();
          Object.entries(COLUMN_ALIASES_ASSIGNMENTS).forEach(([canonical, aliases]) => {
            const aliasList = Array.isArray(aliases) ? aliases : [aliases];
            const matches = aliasList.concat([canonical]).some(a => a && String(a).trim().toLowerCase() === hLower);
            if (matches) headerIndexMap[canonical] = idx;
          });
        });

        // Helper: find first property in entry whose key matches predicate (case-insensitive)
        const findEntryProp = (entry, predicate) => {
          for (const k of Object.keys(entry)) {
            if (predicate(k) && entry[k] !== undefined && entry[k] !== null && String(entry[k]).trim() !== '') {
              return entry[k];
            }
          }
          return null;
        };

        // Build processed assignment rows using headerIndexMap so only desired columns are included
        assignmentsRows.forEach(row => {
          const entry = {};
          Object.entries(headerIndexMap).forEach(([canonical, idx]) => {
            entry[canonical] = row[idx];
          });

          // ensure a default Submission flag for testing (don't overwrite existing value)
          if (entry.Submission === undefined) entry.Submission = false;
          if (entry.submission === undefined) entry.submission = false;

          // Locate student name (any key containing 'name'), id (keys containing 'id' or 'identifier'),
          // and gradebook (exact canonical 'Gradebook' or keys containing 'gradebook')
          const nameVal = findEntryProp(entry, k => k.toLowerCase().includes('name'));
          const idVal = findEntryProp(entry, k => k.toLowerCase().includes('id') || k.toLowerCase().includes('identifier'));
          const gradebookVal = (entry.Gradebook && String(entry.Gradebook).trim() !== '') ? entry.Gradebook : findEntryProp(entry, k => k.toLowerCase().includes('gradebook'));

          if (!nameVal && !idVal && !gradebookVal) return; // nothing to group by

          // push under Gradebook key if available (use string form)
          if (gradebookVal) {
            const gbKey = String(gradebookVal);
            if (!assignmentsMap[gbKey]) assignmentsMap[gbKey] = [];
            assignmentsMap[gbKey].push(entry);
          }

          // push under StudentName key if available
          if (nameVal) {
            if (!assignmentsMap[nameVal]) assignmentsMap[nameVal] = [];
            assignmentsMap[nameVal].push(entry);
          }

          // push under ID key if available
          if (idVal !== null && idVal !== undefined) {
            const idKey = String(idVal);
            if (!assignmentsMap[idKey]) assignmentsMap[idKey] = [];
            assignmentsMap[idKey].push(entry);
          }
        });
      }

      // --- Process history sheet (using COLUMN_ALIASES_HISTORY to decide which columns to include) ---
      let historyMap = {};
      if (historyRange.values && historyRange.values.length > 1) {
        const historyHeaders = historyRange.values[0].map(h => (typeof h === 'string' ? h.trim() : h));
        const historyRows = historyRange.values.slice(1);

        // Build headerIndexMap: canonicalKey -> column index
        const headerIndexMap = {};
        historyHeaders.forEach((header, idx) => {
          if (!header) return;
          const hLower = String(header).trim().toLowerCase();
          Object.entries(COLUMN_ALIASES_HISTORY).forEach(([canonical, aliases]) => {
            // allow aliases to be array or single string
            const aliasList = Array.isArray(aliases) ? aliases : [aliases];
            // check aliases and also the canonical name itself
            const matches = aliasList.concat([canonical]).some(a => a && String(a).trim().toLowerCase() === hLower);
            if (matches) headerIndexMap[canonical] = idx;
          });
        });

        // Build processed history rows using headerIndexMap so only desired columns are included
        historyRows.forEach(row => {
          const entry = {};
          Object.entries(headerIndexMap).forEach(([canonical, idx]) => {
            entry[canonical] = row[idx];
          });

          const id =
            entry.StudentID ||
            entry.ID ||
            entry['Student ID'] ||
            entry['Student identifier'] ||
            "";
          if (!id) return;
          const key = String(id).toLowerCase();
          if (!historyMap[key]) historyMap[key] = [];
          historyMap[key].push(entry);
        });

        // Log the built historyMap keys (counts) for debugging
        try {
          safeLog('Built historyMap keys:', Object.keys(historyMap).length);
        } catch (e) {
          /* ignore logging errors */
        }
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
          // Attach assignments for this student (prefer Gradebook, then StudentName, then ID, fallback to case-insensitive name)
          // prefer exact Gradebook match first, then exact StudentName, then ID, then case-insensitive name
          try {
            let assignmentsForStudent = [];
            // 1) Gradebook match
            const gradebookKey = studentInfo.Gradebook ? String(studentInfo.Gradebook) : "";
            if (gradebookKey && assignmentsMap[gradebookKey]) {
              assignmentsForStudent = assignmentsMap[gradebookKey];
            } else {
              // 2) exact StudentName
              const nameKey = studentInfo.StudentName || "";
              if (nameKey && assignmentsMap[nameKey]) {
                assignmentsForStudent = assignmentsMap[nameKey];
              // 3) exact ID
              } else if (id && assignmentsMap[id]) {
                assignmentsForStudent = assignmentsMap[id];
              } else if (nameKey) {
                // 4) case-insensitive name match
                const lowerName = String(nameKey).trim().toLowerCase();
                for (const k of Object.keys(assignmentsMap)) {
                  if (String(k).trim().toLowerCase() === lowerName) {
                    assignmentsForStudent = assignmentsMap[k];
                    break;
                  }
                }
              }
            }
            studentInfo.Assignments = Array.isArray(assignmentsForStudent) ? assignmentsForStudent : [];
          } catch (e) {
            studentInfo.Assignments = [];
          }
          // Log this student's history array for debugging
          try {
            //safeLog('History for', studentInfo.StudentName || id, studentInfo.History);
          } catch (e) {
            /* ignore logging errors */
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

    // Notify any global refresh handler and log the result for debugging
    try {
      console.log('loadCache returned (wrapper):', result);
      if (typeof window !== 'undefined' && typeof window.refreshCache === 'function') {
        // call without awaiting to avoid blocking callers
        try { window.refreshCache(result); } catch (_) {}
      }
    } catch (_) {}

    return result;
  } catch (error) {
    console.error('loadCache error', error);
    return { status: 'error', message: 'An error occurred while loading the data. Please try again.', data: null, studentCache: [], studentRowIndices: [], headers: [], assignmentsMap: {} };
  }
}
