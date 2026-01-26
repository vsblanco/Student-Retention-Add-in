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
const FORMAT_BATCH_SIZE = 100;

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
 * Helper to generate the Retention Outreach Message.
 * Centralized logic for LDA tags, Missing Assignments, and Explicit DNC.
 * @param {string} sId - Student ID
 * @param {Map} ldaMap - The map of LDA data
 * @param {number} missingVal - The value from "Missing Assignments" column
 * @param {string} tableContext - 'LDA_Table' or 'Failing_Table'
 * @param {Map} dncMap - Map of ID -> Tag Text
 * @returns {string|null} - The formatted message or null
 */
function getRetentionMessage(sId, ldaMap, missingVal, tableContext, dncMap) {
    // Priority 1: Explicit DNC (Highest Priority - Stop everything)
    if (sId && dncMap.has(sId)) {
        const dncTag = dncMap.get(sId);
        // Check if tag is "dnc" by itself (trimmed and lowered)
        if (dncTag && dncTag.trim() === 'dnc') {
            return "[Retention] DNC";
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
            const mm = String(ldaObj.date.getMonth() + 1).padStart(2, '0');
            const dd = String(ldaObj.date.getDate()).padStart(2, '0');
            const yy = String(ldaObj.date.getFullYear()).slice(-2);
            return `[Retention] Student will engage on ${mm}-${dd}-${yy}`;
        }
    }

    return null;
}

/**
 * Main function to create the LDA report.
 * @param {Object} userOverrides - The settings from the UI (daysOut, includeFailingList, etc.)
 * @param {Function} onProgress - Callback to update UI steps: (stepId, status) => void
 * @param {Function} onBatchProgress - Optional callback for batch progress: (current, total, phase, tableName) => void
 */
export async function createLDA(userOverrides, onProgress, onBatchProgress = null) {
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

            // Read cell colors from a sample (first 500 rows) to build color map
            // We don't need all rows - just enough to capture the value->color mappings
            const colorSampleRows = Math.min(500, totalRows);
            const colorSampleRange = masterSheet.getRangeByIndexes(0, 0, colorSampleRows, totalCols);
            const colorSampleProps = colorSampleRange.getCellProperties({
                format: { fill: { color: true } }
            });

            let historyData = null;
            const hasHistory = sheets.items.some(s => s.name === SHEET_NAMES.HISTORY);
            if (hasHistory) {
                const histSheet = sheets.getItem(SHEET_NAMES.HISTORY);
                const histRange = histSheet.getUsedRange();
                histRange.load("values");
                historyData = histRange;
            }

            await context.sync();

            const masterColors = colorSampleProps.value;
            
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

            // --- NEW: Implicit Column Detection ---
            // 1. Mark which Master List indices are covered by User Settings
            const usedMasterIndices = new Set();
            settings.columns.forEach(col => {
                const idx = getColIndex(col.name);
                if (idx !== -1) usedMasterIndices.add(idx);
            });

            // 2. Find unused Master List headers and append them as hidden
            const outputColumns = [...settings.columns]; // Start with user config
            
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


            // --- STEP 5: Creating Sheet ---
            if (onProgress) onProgress('createSheet', 'active');

            const today = new Date();
            const dateStr = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
            const baseName = `LDA ${dateStr}`;
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
                        
                        // Regex to find "LDA" followed by a Date (e.g., LDA 10/25/25, Contacted LDA 12-01-25)
                        // \blda\b : Matches "lda" as a whole word (avoids matching "folder")
                        // .*?     : Non-greedy match for any chars in between (spaces, colons, words)
                        // (...)   : Captures the date part (digits, separator, digits, separator, digits)
                        const ldaRegex = /\blda\b.*?(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;

                        for (let i = hValues.length - 1; i > 0; i--) {
                            const hid = hValues[i][hIdIdx];
                            const htagRaw = String(hValues[i][hTagIdx] || '');
                            const htagLower = htagRaw.toLowerCase().trim();
                            
                            // Check for DNC (Only if toggle is ON)
                            if (settings.includeDNCTag && hid && htagLower.includes('dnc')) {
                                dncMap.set(hid, htagLower);
                            }

                            // Check for LDA (Improved Regex Logic)
                            if (settings.includeLDATag && hid && !ldaFollowUpMap.has(hid)) {
                                const match = htagRaw.match(ldaRegex); // Match against raw or lower, regex has /i flag

                                if (match) {
                                    // match[1] contains the captured date string
                                    const dateString = match[1];
                                    const ldaDate = new Date(dateString);

                                    // Valid Date check & Future Date check
                                    if (!isNaN(ldaDate.getTime())) {
                                        ldaDate.setHours(0,0,0,0);
                                        // Include Today or Future dates
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

                // 2. Generate Retention Message using helper
                const retentionMsg = getRetentionMessage(sId, ldaFollowUpMap, missingVal, tableContext, dncMap);
                
                // 3. Determine Highlighting Logic
                const isLda = sId && ldaFollowUpMap.has(sId);
                const isRetentionActive = !!retentionMsg;
                
                // Determine Row/Partial Color:
                let partialRowColor = "#FFEDD5"; // Orange Default
                if (retentionMsg && retentionMsg.includes("DNC")) {
                    partialRowColor = "#FFC7CE"; // Red for DNC
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
            outputColumns.forEach((colConfig, idx) => {
                if (colConfig.hidden) {
                    newSheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn().columnHidden = true;
                }
            });

            if (onProgress) onProgress('format', 'completed');

            // --- STEP 8: Finalize ---
            if (onProgress) onProgress('finalize', 'active');
            await context.sync();
            if (onProgress) onProgress('finalize', 'completed');

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
    await context.sync();

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
        await context.sync();
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
        for (const op of cellColorOps) {
            const range = bodyRange.getRangeByIndexes(op.rowIdx, op.startCol, 1, op.endCol - op.startCol + 1);
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
        await context.sync();

        formatBatch++;
        if (onBatchProgress) {
            onBatchProgress(formatBatch, totalFormatBatches, 'formatting');
        }
    }
}