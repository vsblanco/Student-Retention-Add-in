/*
 * Timestamp: 2026-01-26 00:00:00
 * Version: 2.22.0
 * Author: Gemini (for Victor)
 * Description: Core logic for creating LDA reports.
 * Update: Optimize formatting performance by using range-based operations instead of cell-by-cell.
 *         Consecutive cells with the same color are merged into single range operations, reducing
 *         API calls from thousands to hundreds per batch. This dramatically speeds up formatting.
 */

// Hardcoded sheet names (unless these are also settings, usually they are static)
const SHEET_NAMES = {
    MASTER_LIST: "Master List",
    HISTORY: "Student History"
};

// Batch size for chunked operations to avoid payload size limits
// Excel Add-ins have ~5MB payload limits; 500 rows is a safe batch size for data
const BATCH_SIZE = 500;

// Smaller batch size for formatting operations (colors, fonts, formulas per cell)
// These are more expensive as each cell operation adds to the request queue
// Kept small (50) to stay well under Excel's per-sync operation limits,
// especially with partial-row cell highlights generating multiple range calls per row.
const FORMAT_BATCH_SIZE = 50;

/**
 * Helper to convert Excel serial date to MM-DD-YY string
 */
function formatExcelDate(serial) {
    if (!serial || isNaN(serial)) return serial;
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    const utcDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60000);
    const mm = String(utcDate.getMonth() + 1).padStart(2, '0');
    const dd = String(utcDate.getDate()).padStart(2, '0');
    const yy = String(utcDate.getFullYear()).slice(-2);
    return `${mm}-${dd}-${yy}`;
}

/**
 * Helper to format a date value into a friendly relative day string.
 * - "today" / "tomorrow" for 0/1 days out
 * - "this [Day]" for 2-6 days out
 * - "next [Day]" for 7-13 days out
 * - Falls back to "Month Dayth" (e.g. March 14th) for anything else
 */
function formatFriendlyDate(value) {
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

    // Check if the value is already "Today" or "Tomorrow" (case-insensitive)
    if (typeof value === 'string') {
        const lower = value.trim().toLowerCase();
        if (lower === 'today') return 'today';
        if (lower === 'tomorrow') return 'tomorrow';
    }

    // Try to parse into a Date object
    let targetDate = null;

    if (typeof value === 'number') {
        // Excel serial number
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        targetDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60000);
    } else if (typeof value === 'string') {
        // Try MM-DD-YY format
        const match = value.trim().match(/^(\d{2})-(\d{2})-(\d{2})$/);
        if (match) {
            const year = 2000 + parseInt(match[3], 10);
            const month = parseInt(match[1], 10) - 1;
            const day = parseInt(match[2], 10);
            targetDate = new Date(year, month, day);
        }
    }

    if (!targetDate || isNaN(targetDate.getTime())) {
        // Can't parse, return as-is
        return typeof value === 'number' ? formatExcelDate(value) : value;
    }

    // Compare with today (strip time)
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    targetDate.setHours(0, 0, 0, 0);

    const diffDays = Math.round((targetDate - today) / (1000 * 60 * 60 * 24));

    if (diffDays === 0) return 'today';
    if (diffDays === 1) return 'tomorrow';
    if (diffDays >= 2 && diffDays <= 6) return `this ${dayNames[targetDate.getDay()]}`;
    if (diffDays >= 7 && diffDays <= 13) return `next ${dayNames[targetDate.getDay()]}`;

    // Beyond 13 days or in the past, use "Month Dayth" format
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    const day = targetDate.getDate();
    const suffix = (day === 1 || day === 21 || day === 31) ? 'st'
        : (day === 2 || day === 22) ? 'nd'
        : (day === 3 || day === 23) ? 'rd' : 'th';
    return `${monthNames[targetDate.getMonth()]} ${day}${suffix}`;
}

/**
 * Helper to generate the Retention Outreach Message.
 * Centralized logic for LDA tags, Missing Assignments, and Explicit DNC.
 * @param {string} sId - Student ID
 * @param {Map} ldaMap - The map of LDA data
 * @param {number} missingVal - The value from "Missing Assignments" column
 * @param {string} tableContext - 'LDA_Table' or 'Failing_Table'
 * @param {Map} dncMap - Map of ID -> Tag Text
 * @returns {string|null} - The formatted message or null
 */
function getRetentionMessage(sId, ldaMap, missingVal, tableContext, dncMap, nextAssignmentDueVal) {
    // Priority 1: Explicit DNC (Highest Priority - Stop everything)
    if (sId && dncMap.has(sId)) {
        const dncTag = dncMap.get(sId);
        if (dncTag) {
            // Split comma-separated tags and check each individually
            const individualTags = dncTag.split(',').map(t => t.trim());
            const hasExcludableDnc = individualTags.some(tag =>
                tag.includes('dnc') && tag !== 'dnc - phone' && tag !== 'dnc - other phone'
            );
            if (hasExcludableDnc) {
                return "Do not contact";
            }
        }
    }

    // Priority 2: Waiting for Grade (Strictly for Failing Table)
    //if (tableContext === 'Failing_Table' && typeof missingVal === 'number' && missingVal === 0) {
    //    return "[Retention] Waiting for Prof. to grade";
    //}

    // Priority 3: LDA Tag (Applies to any table)
    if (sId && ldaMap.has(sId)) {
        const ldaObj = ldaMap.get(sId);
        if (ldaObj && ldaObj.date) {
            const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
            const day = ldaObj.date.getDate();
            const suffix = (day === 1 || day === 21 || day === 31) ? 'st'
                : (day === 2 || day === 22) ? 'nd'
                : (day === 3 || day === 23) ? 'rd' : 'th';
            const friendlyDate = `${monthNames[ldaObj.date.getMonth()]} ${day}${suffix}`;
            return `[Follow up] Student will engage on ${friendlyDate}`;
        }
    }

    // Priority 4: Zero missing assignments with a next assignment due date
    if (typeof missingVal === 'number' && missingVal === 0 && nextAssignmentDueVal) {
        const formattedDate = formatFriendlyDate(nextAssignmentDueVal);
        return `Student's next assignment is due ${formattedDate}.`;
    }

    return null;
}

/**
 * Main function to create the LDA report.
 * @param {Object} userOverrides - The settings from the UI (daysOut, includeFailingList, etc.)
 * @param {Function} onProgress - Callback to update UI steps: (stepId, status) => void
 * @param {Function} onBatchProgress - Optional callback for batch progress: (current, total, phase, tableName) => void
 * @param {Function} onCampusProgress - Optional callback for multi-campus progress: (campusName, campusIndex, totalCampuses, status) => void
 */
export async function createLDA(userOverrides, onProgress, onBatchProgress = null, onCampusProgress = null) {
    try {
        // --- STEP 1: Validate Settings & Environment ---
        if (onProgress) onProgress('validate', 'active');
        
        let workbookSettings = {};
        if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
             const settings = Office.context.document.settings.get('workbookSettings');
             if (settings && typeof settings === 'object') {
                 workbookSettings = settings;
             }
        }

        const settings = {
            daysOut: userOverrides.daysOut ?? 5,
            includeFailingList: userOverrides.includeFailingList ?? false,
            includeLDATag: userOverrides.includeLDATag ?? true,
            includeDNCTag: userOverrides.includeDNCTag ?? true,
            sheetNameMode: userOverrides.sheetNameMode ?? 'date',
            columns: workbookSettings.columns || []
        };

        if (settings.columns.length === 0) {
            throw new Error("No column settings found. Please configure columns in the Settings tab first.");
        }

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) {
                throw new Error(`Workbook missing "${SHEET_NAMES.MASTER_LIST}" sheet.`);
            }
            if (onProgress) onProgress('validate', 'completed');

            // --- STEP 2: Read Master List & Scan Colors ---
            if (onProgress) onProgress('read', 'active');

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();

            // First, get dimensions only
            masterRange.load("rowCount, columnCount");
            await context.sync();

            const totalRows = masterRange.rowCount;
            const totalCols = masterRange.columnCount;

            // Read values and formulas in batches to avoid response payload limits
            let masterValues = [];
            let masterFormulas = [];

            for (let startRow = 0; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = masterSheet.getRangeByIndexes(startRow, 0, rowsToRead, totalCols);
                batchRange.load("values, formulas");
                await context.sync();

                masterValues = masterValues.concat(batchRange.values);
                masterFormulas = masterFormulas.concat(batchRange.formulas);

                // Report progress for large datasets
                if (onBatchProgress && totalRows > BATCH_SIZE) {
                    const currentBatch = Math.floor(startRow / BATCH_SIZE) + 1;
                    const totalBatches = Math.ceil(totalRows / BATCH_SIZE);
                    onBatchProgress(currentBatch, totalBatches, 'reading', 'Master_List');
                }
            }

            const headers = masterValues[0];

            // Read cell colors from a sample to build color map
            // getCellProperties returns ~100-200 bytes per cell; cap total cells to stay under ~5MB payload limit
            const MAX_COLOR_CELLS = 12000;
            const colorSampleRows = Math.min(
                Math.max(50, Math.floor(MAX_COLOR_CELLS / totalCols)),
                500,
                totalRows
            );
            const colorSampleRange = masterSheet.getRangeByIndexes(0, 0, colorSampleRows, totalCols);
            let colorSampleProps;
            try {
                colorSampleProps = colorSampleRange.getCellProperties({
                    format: { fill: { color: true } }
                });
            } catch (e) {
                console.warn('getCellProperties failed, skipping color mapping', e);
                colorSampleProps = null;
            }

            let historyData = null;
            const hasHistory = sheets.items.some(s => s.name === SHEET_NAMES.HISTORY);
            if (hasHistory) {
                const histSheet = sheets.getItem(SHEET_NAMES.HISTORY);
                const histRange = histSheet.getUsedRange();
                histRange.load("values");
                historyData = histRange;
            }

            await context.sync();

            // getCellProperties may have failed (null) or may not have resolved
            let masterColors = null;
            try {
                if (colorSampleProps && colorSampleProps.value) {
                    masterColors = colorSampleProps.value;
                }
            } catch (e) {
                console.warn('Color sample resolution failed, skipping color mapping', e);
                masterColors = null;
            }
            
            // --- UPDATED: Space-Insensitive Matching ---
            const getColIndex = (settingName) => {
                // Helpers: clean (trim/lower) and strip (remove all spaces)
                const cleanStr = (s) => String(s || '').trim().toLowerCase();
                const stripStr = (s) => cleanStr(s).replace(/\s+/g, '');

                const targetNameStripped = stripStr(settingName);

                // 1. Find Config
                const colConfig = settings.columns.find(c => 
                    stripStr(c.name) === targetNameStripped
                );
                // If checking for a raw header that isn't in settings config but passed as a name
                let aliases = [];
                if (colConfig) {
                    if (Array.isArray(colConfig.alias)) {
                        aliases = colConfig.alias;
                    } else if (colConfig.alias) {
                        aliases = [colConfig.alias];
                    }
                }
                
                const candidates = [settingName, ...aliases]; // settingName might be a raw header now

                // 2. Find match in Excel Headers
                for (const rawCand of candidates) {
                    const candStripped = stripStr(rawCand);
                    const idx = headers.findIndex(h => stripStr(h) === candStripped);
                    if (idx !== -1) return idx;
                }

                return -1;
            };

            // --- Implicit Column Detection ---
            // 1. Filter settings columns to only those found in the Master List (skip missing)
            const matchedColumns = settings.columns.filter(col => getColIndex(col.name) !== -1);

            // 2. Mark which Master List indices are covered by matched settings columns
            const usedMasterIndices = new Set();
            matchedColumns.forEach(col => {
                const idx = getColIndex(col.name);
                if (idx !== -1) usedMasterIndices.add(idx);
            });

            // 3. Find unused Master List headers and append them as hidden
            const outputColumns = [...matchedColumns]; // Start with matched user config

            headers.forEach((h, idx) => {
                if (!usedMasterIndices.has(idx) && h && String(h).trim() !== "") {
                    // This is a column in Master List not in settings.
                    // Add it, but hide it.
                    outputColumns.push({
                        name: h,
                        hidden: true
                    });
                }
            });

            // --- Retrieve Key Indices ---
            const daysOutIdx = getColIndex('Days Out');
            const gradeIdx = getColIndex('Grade');
            const studentIdIdx = getColIndex('Student Number');
            
            // Look for "Missing Assignments" column
            let missingIdx = getColIndex('Missing Assignments');
            if (missingIdx === -1) {
                 // Fallback scan
                 missingIdx = headers.findIndex(h => String(h).trim().toLowerCase().includes('missing'));
            }

            // Look for "Next Assignment Due" column
            let nextAssignmentDueIdx = getColIndex('Next Assignment Due');
            if (nextAssignmentDueIdx === -1) {
                nextAssignmentDueIdx = headers.findIndex(h =>
                    String(h).trim().toLowerCase().replace(/\s+/g, '') === 'nextassignmentdue'
                );
            }

            if (daysOutIdx === -1) throw new Error("Could not find 'Days Out' column in Master List. Check Settings.");

            if (onProgress) onProgress('read', 'completed');


            // --- STEP 2b: Build Color Map (Value -> Color) ---
            const columnColorMaps = new Map();

            // Re-sort outputColumns to ensure processing order (though less critical now)
            // We only really care about mapping indices for color scanning
            
            const outputColIndices = outputColumns
                .map(c => getColIndex(c.name))
                .filter(idx => idx !== -1);

            // --- EXCLUDED COLORS SET ---
            const EXCLUDED_COLORS = new Set([
                '#ffffff', // White
                '#add8e6'  // LightBlue (Standard Excel/HTML name)
            ]);

            // Only iterate over the sampled color rows (masterColors has limited rows)
            if (masterColors) {
                const colorRowLimit = Math.min(masterColors.length, masterValues.length);
                for (let r = 1; r < colorRowLimit; r++) {
                    outputColIndices.forEach(cIdx => {
                        const val = masterValues[r][cIdx];
                        const rawColor = masterColors[r]?.[cIdx]?.format?.fill?.color;
                        const normColor = String(rawColor || '').toLowerCase();

                        // Check if value exists, color exists, and color is NOT in the excluded list
                        if (val && rawColor && !EXCLUDED_COLORS.has(normColor)) {
                            if (!columnColorMaps.has(cIdx)) {
                                columnColorMaps.set(cIdx, new Map());
                            }
                            columnColorMaps.get(cIdx).set(String(val), rawColor);
                        }
                    });
                }
            }

            // --- STEP 3: Filtering by Days Out ---
            if (onProgress) onProgress('filter', 'active');

            const dataRows = [];
            for (let i = 1; i < masterValues.length; i++) {
                const daysOutVal = masterValues[i][daysOutIdx];
                if (typeof daysOutVal === 'number' && daysOutVal >= settings.daysOut) {
                    dataRows.push({
                        values: masterValues[i],
                        formulas: masterFormulas[i],
                        originalIndex: i
                    });
                }
            }
            
            dataRows.sort((a, b) => (b.values[daysOutIdx] || 0) - (a.values[daysOutIdx] || 0));
            
            if (onProgress) onProgress('filter', 'completed');

            // --- STEP 4: Filtering by Grades (Failing) ---
            if (onProgress) onProgress('failing', 'active');
            
            let failingRows = [];
            if (settings.includeFailingList && gradeIdx !== -1) {
                for (let i = 1; i < masterValues.length; i++) {
                    const gradeVal = masterValues[i][gradeIdx];
                    const daysOutVal = masterValues[i][daysOutIdx];
                    
                    const isFailing = (typeof gradeVal === 'number') && (gradeVal < 0.60 || (gradeVal >= 1 && gradeVal < 60));
                    const isRecent = (typeof daysOutVal === 'number') && (daysOutVal <= 4);

                    if (isFailing && isRecent) {
                        failingRows.push({
                            values: masterValues[i],
                            formulas: masterFormulas[i],
                            originalIndex: i
                        });
                    }
                }
                failingRows.sort((a, b) => (a.values[gradeIdx] || 0) - (b.values[gradeIdx] || 0));
            }
            if (onProgress) onProgress('failing', 'completed');


            // --- Detect Multi-Campus Mode ---
            let isMultiCampus = false;
            let campusList = [];
            const campusIdx = getColIndex('Campus');

            if (settings.sheetNameMode === 'campus' && campusIdx !== -1) {
                const campusSet = new Set();
                for (let i = 1; i < masterValues.length; i++) {
                    const val = String(masterValues[i][campusIdx] || '').trim();
                    if (val) campusSet.add(val);
                }
                campusList = Array.from(campusSet).sort();
                isMultiCampus = campusList.length > 1;
            }


            if (isMultiCampus) {
                // ============================================================
                // MULTI-CAMPUS FLOW: Create one LDA sheet per campus
                // ============================================================

                // --- STEP 5 (Tags - global): Applying LDA & DNC Tags ---
                if (onProgress) onProgress('tags', 'active');

                const dncMap = new Map();
                const ldaFollowUpMap = new Map();

                if (historyData && studentIdIdx !== -1) {
                    const hValues = historyData.values;
                    if (hValues.length > 0) {
                        const hHeaders = hValues[0].map(h => String(h).toLowerCase().trim());
                        const hIdIdx = hHeaders.findIndex(h => h.includes('student') && h.includes('id') || h.includes('number'));
                        const hTagIdx = hHeaders.indexOf('tag');

                        if (hIdIdx !== -1 && hTagIdx !== -1) {
                            const todayTime = new Date().setHours(0,0,0,0);
                            const ldaRegex = /\blda\b.*?(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;

                            for (let i = hValues.length - 1; i > 0; i--) {
                                const hid = hValues[i][hIdIdx];
                                const htagRaw = String(hValues[i][hTagIdx] || '');
                                const htagLower = htagRaw.toLowerCase().trim();

                                if (settings.includeDNCTag && hid && htagLower.includes('dnc')) {
                                    dncMap.set(hid, htagLower);
                                }

                                if (settings.includeLDATag && hid && !ldaFollowUpMap.has(hid)) {
                                    const match = htagRaw.match(ldaRegex);
                                    if (match) {
                                        const ldaDate = new Date(match[1]);
                                        if (!isNaN(ldaDate.getTime())) {
                                            ldaDate.setHours(0,0,0,0);
                                            if (ldaDate >= todayTime) {
                                                ldaFollowUpMap.set(hid, { date: ldaDate, text: htagRaw });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (onProgress) onProgress('tags', 'completed');

                // --- Shared format helpers (used by all campus sheets) ---
                const dateColumnIndices = new Set();
                outputColumns.forEach((colConfig) => {
                    const masterIdx = getColIndex(colConfig.name);
                    if (masterIdx === -1) return;
                    const excelHeader = String(masterValues[0][masterIdx] || '').toLowerCase();
                    const isExcluded = /id|no\.|num|code|zip|postal|social|ssn|phone|grade|score|credit|fee|days|count/i.test(excelHeader);
                    if (isExcluded) return;
                    let dateCount = 0;
                    let numCount = 0;
                    const limit = Math.min(masterValues.length, 100);
                    for (let i = 1; i < limit; i++) {
                        const val = masterValues[i][masterIdx];
                        if (typeof val === 'number') {
                            numCount++;
                            if (val > 10958 && val < 73051) dateCount++;
                        }
                    }
                    if (numCount > 0 && (dateCount / numCount) > 0.5) {
                        dateColumnIndices.add(colConfig.name);
                    }
                });

                const outreachColIndex = outputColumns.findIndex(c => c.name === 'Outreach');

                const buildOutputRow = (rowObj, tableContext) => {
                    const cells = [];
                    const formulas = [];
                    let rowColor = null;
                    let cellHighlights = [];
                    const sId = rowObj.values[studentIdIdx];
                    const missingVal = (missingIdx !== -1) ? rowObj.values[missingIdx] : null;
                    const nextAssignmentDueVal = (nextAssignmentDueIdx !== -1) ? rowObj.values[nextAssignmentDueIdx] : null;
                    const retentionMsg = getRetentionMessage(sId, ldaFollowUpMap, missingVal, tableContext, dncMap, nextAssignmentDueVal);
                    const isRetentionActive = !!retentionMsg;
                    const isNextAssignmentDue = retentionMsg && retentionMsg.startsWith("Student's next assignment is due");
                    let partialRowColor = "#FFEDD5";
                    if (retentionMsg && retentionMsg.includes("DNC")) {
                        partialRowColor = "#FFC7CE";
                    } else if (isNextAssignmentDue) {
                        partialRowColor = "#e2efda";
                    }
                    if (isRetentionActive && outreachColIndex === -1) {
                        rowColor = partialRowColor;
                    }
                    outputColumns.forEach((colConfig, colOutIdx) => {
                        const masterIdx = getColIndex(colConfig.name);
                        let val = (masterIdx !== -1) ? rowObj.values[masterIdx] : "";
                        let form = (masterIdx !== -1) ? rowObj.formulas[masterIdx] : null;
                        if (colConfig.name === 'Gradebook' && form && String(form).startsWith('=HYPERLINK')) {
                            // Keep formula
                        } else if (colConfig.name === 'Gradebook' && val && String(val).startsWith('http')) {
                            form = `=HYPERLINK("${val}", "Link")`;
                            val = "Link";
                        }
                        if (isRetentionActive && outreachColIndex !== -1 && colOutIdx <= outreachColIndex) {
                            cellHighlights.push({ colIndex: colOutIdx, color: partialRowColor });
                        }
                        if (masterIdx !== -1 && val) {
                            const colMap = columnColorMaps.get(masterIdx);
                            if (colMap && colMap.has(String(val))) {
                                cellHighlights.push({ colIndex: colOutIdx, color: colMap.get(String(val)) });
                            }
                        }
                        if (colConfig.name === 'Outreach' && retentionMsg) {
                            val = retentionMsg;
                        }
                        if (settings.includeDNCTag && dncMap.has(sId)) {
                            if (colConfig.name === 'Phone' || colConfig.name === 'Other Phone') {
                                cellHighlights.push({ colIndex: colOutIdx, color: "#FFC7CE", strikethrough: true });
                            }
                        }
                        cells.push(val);
                        formulas.push(form);
                    });
                    return { cells, formulas, rowColor, cellHighlights };
                };

                // --- Process each campus ---
                if (onProgress) onProgress('createSheet', 'active');

                for (let ci = 0; ci < campusList.length; ci++) {
                    const campusName = campusList[ci];
                    if (onCampusProgress) onCampusProgress(campusName, ci, campusList.length, 'active');

                    // Filter data for this campus
                    const campusDataRows = dataRows.filter(r =>
                        String(r.values[campusIdx] || '').trim() === campusName
                    );
                    const campusFailingRows = failingRows.filter(r =>
                        String(r.values[campusIdx] || '').trim() === campusName
                    );

                    // Create sheet for this campus
                    let sheetName = campusName;
                    let counter = 2;
                    sheets.load("items/name");
                    await context.sync();
                    const existingNames = sheets.items.map(s => s.name);
                    while (existingNames.includes(sheetName)) {
                        sheetName = `${campusName} (${counter++})`;
                    }

                    const newSheet = sheets.add(sheetName);
                    if (ci === 0) newSheet.activate();
                    await context.sync();

                    // Write LDA table for this campus
                    if (campusDataRows.length > 0) {
                        await writeTable(
                            context, newSheet, 0, `LDA_${ci}`,
                            outputColumns,
                            campusDataRows.map(r => buildOutputRow(r, 'LDA_Table')),
                            masterSheet, getColIndex, dateColumnIndices, null
                        );
                    } else {
                        const headerRange = newSheet.getRangeByIndexes(0, 0, 1, outputColumns.length);
                        headerRange.values = [outputColumns.map(c => c.name)];
                    }

                    // Write Failing table for this campus (if applicable)
                    let nextRow = campusDataRows.length + 4;
                    if (settings.includeFailingList && campusFailingRows.length > 0) {
                        const title = newSheet.getRangeByIndexes(nextRow - 1, 0, 1, 1);
                        title.values = [["Failing Students (Active)"]];
                        title.format.font.bold = true;

                        await writeTable(
                            context, newSheet, nextRow, `Failing_${ci}`,
                            outputColumns,
                            campusFailingRows.map(r => buildOutputRow(r, 'Failing_Table')),
                            masterSheet, getColIndex, dateColumnIndices, null
                        );
                    }

                    // Autofit & hide columns
                    newSheet.getUsedRange().getEntireColumn().format.autofitColumns();
                    await context.sync();

                    // Batch hidden column operations to avoid queue overflow
                    const campusHiddenIndices = outputColumns
                        .map((colConfig, idx) => colConfig.hidden ? idx : -1)
                        .filter(idx => idx !== -1);
                    const CAMPUS_HIDE_BATCH = 50;
                    for (let hi = 0; hi < campusHiddenIndices.length; hi += CAMPUS_HIDE_BATCH) {
                        const batch = campusHiddenIndices.slice(hi, hi + CAMPUS_HIDE_BATCH);
                        batch.forEach(idx => {
                            newSheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn().columnHidden = true;
                        });
                        await context.sync();
                    }
                    if (onCampusProgress) onCampusProgress(campusName, ci, campusList.length, 'completed');
                }

                if (onProgress) onProgress('createSheet', 'completed');
                if (onProgress) onProgress('format', 'completed');
                if (onProgress) onProgress('finalize', 'completed');

            } else {
                // ============================================================
                // SINGLE / DATE FLOW (original behavior)
                // ============================================================

                // --- STEP 5: Creating Sheet ---
                if (onProgress) onProgress('createSheet', 'active');

                let baseName;
                if (settings.sheetNameMode === 'campus') {
                    let campusName = '';
                    if (campusIdx !== -1) {
                        for (let i = 1; i < masterValues.length; i++) {
                            const val = String(masterValues[i][campusIdx] || '').trim();
                            if (val) { campusName = val; break; }
                        }
                    }
                    baseName = campusName || 'Campus';
                } else {
                    const today = new Date();
                    const dateStr = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
                    baseName = `LDA ${dateStr}`;
                }
                let sheetName = baseName;

                let counter = 2;
                const existingNames = sheets.items.map(s => s.name);
                while (existingNames.includes(sheetName)) {
                    sheetName = `${baseName} (${counter++})`;
                }

                const newSheet = sheets.add(sheetName);
                newSheet.activate();
                await context.sync();
                if (onProgress) onProgress('createSheet', 'completed');


                // --- STEP 6: Applying Tags (Data Processing) ---
                if (onProgress) onProgress('tags', 'active');

                const dncMap = new Map();
                const ldaFollowUpMap = new Map();

                if (historyData && studentIdIdx !== -1) {
                    const hValues = historyData.values;
                    if (hValues.length > 0) {
                        const hHeaders = hValues[0].map(h => String(h).toLowerCase().trim());
                        const hIdIdx = hHeaders.findIndex(h => h.includes('student') && h.includes('id') || h.includes('number'));
                        const hTagIdx = hHeaders.indexOf('tag');

                        if (hIdIdx !== -1 && hTagIdx !== -1) {
                            const todayTime = new Date().setHours(0,0,0,0);
                            const ldaRegex = /\blda\b.*?(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;

                            for (let i = hValues.length - 1; i > 0; i--) {
                                const hid = hValues[i][hIdIdx];
                                const htagRaw = String(hValues[i][hTagIdx] || '');
                                const htagLower = htagRaw.toLowerCase().trim();

                                if (settings.includeDNCTag && hid && htagLower.includes('dnc')) {
                                    dncMap.set(hid, htagLower);
                                }

                                if (settings.includeLDATag && hid && !ldaFollowUpMap.has(hid)) {
                                    const match = htagRaw.match(ldaRegex);
                                    if (match) {
                                        const dateString = match[1];
                                        const ldaDate = new Date(dateString);
                                        if (!isNaN(ldaDate.getTime())) {
                                            ldaDate.setHours(0,0,0,0);
                                            if (ldaDate >= todayTime) {
                                                ldaFollowUpMap.set(hid, { date: ldaDate, text: htagRaw });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (onProgress) onProgress('tags', 'completed');


                // --- STEP 7: Formatting LDA Table ---
                if (onProgress) onProgress('format', 'active');

                // --- 7a. Data-Driven Date Column Detection ---
                const dateColumnIndices = new Set();
                outputColumns.forEach((colConfig) => {
                    const masterIdx = getColIndex(colConfig.name);
                    if (masterIdx === -1) return;

                    const excelHeader = String(masterValues[0][masterIdx] || '').toLowerCase();
                    const isExcluded = /id|no\.|num|code|zip|postal|social|ssn|phone|grade|score|credit|fee|days|count/i.test(excelHeader);
                    if (isExcluded) return;

                    let dateCount = 0;
                    let numCount = 0;
                    const limit = Math.min(masterValues.length, 100);
                    for (let i = 1; i < limit; i++) {
                        const val = masterValues[i][masterIdx];
                        if (typeof val === 'number') {
                            numCount++;
                            if (val > 10958 && val < 73051) {
                                dateCount++;
                            }
                        }
                    }
                    if (numCount > 0 && (dateCount / numCount) > 0.5) {
                        dateColumnIndices.add(colConfig.name);
                    }
                });

                // Determine Outreach Column Index (It might be missing if implicit columns pushed it)
                const outreachColIndex = outputColumns.findIndex(c => c.name === 'Outreach');

                // Row Builder
                const buildOutputRow = (rowObj, tableContext) => {
                    const cells = [];
                    const formulas = [];
                    let rowColor = null;
                    let cellHighlights = [];

                    const sId = rowObj.values[studentIdIdx];

                    // 1. Get critical values
                    const missingVal = (missingIdx !== -1) ? rowObj.values[missingIdx] : null;
                    const nextAssignmentDueVal = (nextAssignmentDueIdx !== -1) ? rowObj.values[nextAssignmentDueIdx] : null;

                    // 2. Generate Retention Message using helper
                    const retentionMsg = getRetentionMessage(sId, ldaFollowUpMap, missingVal, tableContext, dncMap, nextAssignmentDueVal);

                    // 3. Determine Highlighting Logic
                    const isLda = sId && ldaFollowUpMap.has(sId);
                    const isRetentionActive = !!retentionMsg;
                    const isNextAssignmentDue = retentionMsg && retentionMsg.startsWith("Student's next assignment is due");

                    // Determine Row/Partial Color:
                    let partialRowColor = "#FFEDD5"; // Orange Default
                    if (retentionMsg && retentionMsg.includes("DNC")) {
                        partialRowColor = "#FFC7CE"; // Red for DNC
                    } else if (isNextAssignmentDue) {
                        partialRowColor = "#e2efda"; // Light green for zero missing + next assignment due
                    }

                    // Fallback row color if Outreach column is missing
                    if (isRetentionActive && outreachColIndex === -1) {
                         rowColor = partialRowColor;
                    }

                    outputColumns.forEach((colConfig, colOutIdx) => {
                        const masterIdx = getColIndex(colConfig.name);
                        let val = (masterIdx !== -1) ? rowObj.values[masterIdx] : "";
                        let form = (masterIdx !== -1) ? rowObj.formulas[masterIdx] : null;

                        if (colConfig.name === 'Gradebook' && form && String(form).startsWith('=HYPERLINK')) {
                            // Keep formula
                        } else if (colConfig.name === 'Gradebook' && val && String(val).startsWith('http')) {
                            form = `=HYPERLINK("${val}", "Link")`;
                            val = "Link";
                        }

                        // --- Apply Retention Highlight (Partial Row) ---
                        // Only apply up to outreach column if it exists and this column is within range
                        if (isRetentionActive && outreachColIndex !== -1 && colOutIdx <= outreachColIndex) {
                            cellHighlights.push({
                                colIndex: colOutIdx,
                                color: partialRowColor
                            });
                        }

                        // --- Value-Based Color Mapping (Overrides Retention Highlight) ---
                        if (masterIdx !== -1 && val) {
                             const colMap = columnColorMaps.get(masterIdx);
                             if (colMap && colMap.has(String(val))) {
                                 cellHighlights.push({
                                     colIndex: colOutIdx,
                                     color: colMap.get(String(val))
                                 });
                             }
                        }

                        // --- Outreach Message Injection ---
                        if (colConfig.name === 'Outreach' && retentionMsg) {
                            val = retentionMsg;
                        }

                        // --- DNC Highlight (Highest Priority - Phone Columns) ---
                        if (settings.includeDNCTag && dncMap.has(sId)) {
                             if (colConfig.name === 'Phone' || colConfig.name === 'Other Phone') {
                                 cellHighlights.push({
                                     colIndex: colOutIdx,
                                     color: "#FFC7CE",
                                     strikethrough: true
                                 });
                             }
                        }

                        cells.push(val);
                        formulas.push(form);
                    });

                    return { cells, formulas, rowColor, cellHighlights };
                };

                // 7b. Write Main LDA Table
                if (dataRows.length > 0) {
                    // Create a batch progress wrapper that includes table name
                    const ldaBatchProgress = onBatchProgress
                        ? (current, total, phase) => onBatchProgress(current, total, phase, 'LDA_Table')
                        : null;

                    await writeTable(
                        context,
                        newSheet,
                        0,
                        "LDA_Table",
                        outputColumns,
                        dataRows.map(r => buildOutputRow(r, 'LDA_Table')),
                        masterSheet,
                        getColIndex,
                        dateColumnIndices,
                        ldaBatchProgress
                    );
                } else {
                    const headerRange = newSheet.getRangeByIndexes(0, 0, 1, outputColumns.length);
                    headerRange.values = [outputColumns.map(c => c.name)];
                }

                // 7c. Write Failing Table (if applicable)
                let nextRow = dataRows.length + 4;
                if (settings.includeFailingList && failingRows.length > 0) {
                    const title = newSheet.getRangeByIndexes(nextRow - 1, 0, 1, 1);
                    title.values = [["Failing Students (Active)"]];
                    title.format.font.bold = true;

                    // Create a batch progress wrapper that includes table name
                    const failingBatchProgress = onBatchProgress
                        ? (current, total, phase) => onBatchProgress(current, total, phase, 'Failing_Table')
                        : null;

                    await writeTable(
                        context,
                        newSheet,
                        nextRow,
                        "Failing_Table",
                        outputColumns,
                        failingRows.map(r => buildOutputRow(r, 'Failing_Table')),
                        masterSheet,
                        getColIndex,
                        dateColumnIndices,
                        failingBatchProgress
                    );
                }

                // Autofit
                newSheet.getUsedRange().getEntireColumn().format.autofitColumns();

                // --- STEP 7d: Apply Hidden Columns (Must be done LAST after autofit) ---
                // Batch hidden column operations to avoid queue overflow with many columns
                const HIDE_BATCH = 50;
                const hiddenIndices = outputColumns
                    .map((colConfig, idx) => colConfig.hidden ? idx : -1)
                    .filter(idx => idx !== -1);
                for (let i = 0; i < hiddenIndices.length; i += HIDE_BATCH) {
                    const batch = hiddenIndices.slice(i, i + HIDE_BATCH);
                    batch.forEach(idx => {
                        newSheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn().columnHidden = true;
                    });
                    await context.sync();
                }

                if (onProgress) onProgress('format', 'completed');

                // --- STEP 8: Finalize ---
                if (onProgress) onProgress('finalize', 'active');
                await context.sync();
                if (onProgress) onProgress('finalize', 'completed');
            }

        });

    } catch (error) {
        console.error("LDA Generation Error:", error);
        throw error;
    }
}

/**
 * Helper to write a table, COPY Conditional Formatting, and apply custom styles.
 * Uses chunked batch operations to avoid payload size limits with large datasets.
 * @param {Function} onBatchProgress - Optional callback (currentBatch, totalBatches, phase) for progress updates
 */
async function writeTable(context, sheet, startRow, tableName, outputColumns, processedRows, masterSheet, getColIndex, dateColumnNames, onBatchProgress = null) {
    if (processedRows.length === 0) return;

    const rowCount = processedRows.length;
    const colCount = outputColumns.length;
    const headers = outputColumns.map(c => c.name);

    // --- STEP 1: Write headers first ---
    const headerRange = sheet.getRangeByIndexes(startRow, 0, 1, colCount);
    headerRange.values = [headers];
    await context.sync();

    // --- STEP 2: Write data rows in batches to avoid payload size limits ---
    const totalDataBatches = Math.ceil(rowCount / BATCH_SIZE);
    let currentBatch = 0;

    for (let batchStart = 0; batchStart < rowCount; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, rowCount);
        const batchRows = processedRows.slice(batchStart, batchEnd);
        const batchValues = batchRows.map(r => r.cells);

        // Calculate the actual row position in the sheet (startRow + 1 for header + batchStart)
        const batchRange = sheet.getRangeByIndexes(
            startRow + 1 + batchStart,
            0,
            batchValues.length,
            colCount
        );
        batchRange.values = batchValues;

        // Sync after each batch to flush the request and avoid payload limits
        await context.sync();

        currentBatch++;
        if (onBatchProgress) {
            onBatchProgress(currentBatch, totalDataBatches, 'writing');
        }
    }

    // --- STEP 3: Create table after all data is written ---
    const fullRange = sheet.getRangeByIndexes(startRow, 0, rowCount + 1, colCount);
    const table = sheet.tables.add(fullRange, true);
    table.name = tableName + "_" + Math.floor(Math.random() * 1000);
    table.style = "TableStyleLight9";
    try {
        await context.sync();
    } catch (e) {
        console.error(`Table creation sync failed for "${tableName}" (startRow=${startRow}, rows=${rowCount}, cols=${colCount})`, e);
        throw e;
    }

    // --- STEP 4: Copy Conditional Formatting from Master List (in batches) ---
    const cfChecks = [];

    outputColumns.forEach((colConfig, idx) => {
        const masterIdx = getColIndex(colConfig.name);
        if (masterIdx !== -1) {
            const sourceCell = masterSheet.getCell(1, masterIdx);
            cfChecks.push({
                sourceCell,
                targetIndex: idx,
                colName: colConfig.name
            });
        }
    });

    // Process CF in batches to avoid payload limits
    for (let i = 0; i < cfChecks.length; i += BATCH_SIZE) {
        const batch = cfChecks.slice(i, i + BATCH_SIZE);
        batch.forEach(check => {
            const targetColRange = table.columns.getItemAt(check.targetIndex).getDataBodyRange();
            targetColRange.copyFrom(check.sourceCell, Excel.RangeCopyType.formats);
            targetColRange.format.fill.clear();

            if (dateColumnNames.has(check.colName)) {
                targetColRange.numberFormat = [["mm-dd-yy;@"]];
            }
        });
        try {
            await context.sync();
        } catch (e) {
            const colNames = batch.map(c => c.colName).join(', ');
            console.error(`CF copy sync failed on columns: [${colNames}]`, e);
            throw e;
        }
    }

    // --- STEP 5: Apply Custom Row Colors & Cell Highlights (optimized) ---
    // Use range-based operations instead of cell-by-cell for much better performance
    const bodyRange = table.getDataBodyRange();
    const totalFormatBatches = Math.ceil(rowCount / FORMAT_BATCH_SIZE);
    let formatBatch = 0;

    for (let batchStart = 0; batchStart < rowCount; batchStart += FORMAT_BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + FORMAT_BATCH_SIZE, rowCount);

        // Collect all formatting operations for this batch first
        // Then apply them efficiently using range operations

        // 1. Collect row-level colors (apply to full rows)
        const rowColorOps = []; // {rowIdx, color}

        // 2. Collect cell highlights - group consecutive cells with same color per row
        const cellColorOps = []; // {rowIdx, startCol, endCol, color, strikethrough}

        // 3. Collect formulas
        const formulaOps = []; // {rowIdx, colIdx, formula}

        for (let idx = batchStart; idx < batchEnd; idx++) {
            const r = processedRows[idx];

            // Collect row colors
            if (r.rowColor) {
                rowColorOps.push({ rowIdx: idx, color: r.rowColor });
            }

            // Collect and merge consecutive cell highlights with same color
            if (r.cellHighlights.length > 0) {
                // Sort highlights by column index
                const sorted = [...r.cellHighlights].sort((a, b) => a.colIndex - b.colIndex);

                let current = null;
                for (const h of sorted) {
                    if (current && current.color === h.color && current.endCol === h.colIndex - 1 && !h.strikethrough && !current.strikethrough) {
                        // Extend current range
                        current.endCol = h.colIndex;
                    } else {
                        // Start new range
                        if (current) cellColorOps.push(current);
                        current = {
                            rowIdx: idx,
                            startCol: h.colIndex,
                            endCol: h.colIndex,
                            color: h.color,
                            strikethrough: h.strikethrough || false
                        };
                    }
                }
                if (current) cellColorOps.push(current);
            }

            // Collect formulas
            r.formulas.forEach((f, cIdx) => {
                if (f) formulaOps.push({ rowIdx: idx, colIdx: cIdx, formula: f });
            });
        }

        // Apply row colors (these are rare, so individual calls are OK)
        for (const op of rowColorOps) {
            bodyRange.getRow(op.rowIdx).format.fill.color = op.color;
        }

        // Apply cell color ranges (merged ranges = fewer API calls)
        // Use sheet.getRangeByIndexes with absolute positions (bodyRange doesn't have this method)
        // Body starts at startRow + 1 (after header)
        for (const op of cellColorOps) {
            const absoluteRow = startRow + 1 + op.rowIdx;
            const range = sheet.getRangeByIndexes(absoluteRow, op.startCol, 1, op.endCol - op.startCol + 1);
            range.format.fill.color = op.color;
            if (op.strikethrough) {
                range.format.font.strikethrough = true;
                range.format.font.color = "#9C0006";
            }
        }

        // Apply formulas (typically few per batch)
        for (const op of formulaOps) {
            bodyRange.getCell(op.rowIdx, op.colIdx).formulas = [[op.formula]];
        }

        // Sync after each batch
        try {
            await context.sync();
        } catch (e) {
            console.error(`Formatting sync failed on batch ${formatBatch + 1}/${totalFormatBatches} (rows ${batchStart}-${batchEnd - 1}). ` +
                `Ops: ${rowColorOps.length} rowColor, ${cellColorOps.length} cellColor, ${formulaOps.length} formula.`, e);
            throw e;
        }

        formatBatch++;
        if (onBatchProgress) {
            onBatchProgress(formatBatch, totalFormatBatches, 'formatting');
        }
    }
}

/**
 * Detect unique campus values from the Master List.
 * Reads only the Campus column in batches for efficiency.
 * @returns {Promise<string[]>} Array of unique campus names, sorted alphabetically
 */
export async function detectCampuses() {
    try {
        return await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) return [];

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("rowCount, columnCount");
            await context.sync();

            const totalRows = masterRange.rowCount;
            const totalCols = masterRange.columnCount;

            // Read headers only
            const headerRange = masterSheet.getRangeByIndexes(0, 0, 1, totalCols);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0];

            // Find Campus column with case/space insensitive matching
            const campusIdx = headers.findIndex(h => {
                const stripped = String(h || '').trim().toLowerCase().replace(/\s+/g, '');
                return stripped === 'campus';
            });

            if (campusIdx === -1) return [];

            // Read just the campus column in batches
            const campusSet = new Set();
            for (let startRow = 1; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = masterSheet.getRangeByIndexes(startRow, campusIdx, rowsToRead, 1);
                batchRange.load("values");
                await context.sync();

                for (const row of batchRange.values) {
                    const val = String(row[0] || '').trim();
                    if (val) campusSet.add(val);
                }
            }

            return Array.from(campusSet).sort();
        });
    } catch (error) {
        console.error("Campus detection error:", error);
        return [];
    }
}