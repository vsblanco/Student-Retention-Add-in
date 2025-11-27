/*
 * Timestamp: 2025-11-26 22:15:00
 * Version: 2.14.0
 * Author: Gemini (for Victor)
 * Description: Core logic for creating LDA reports.
 * Update: Added Color Exclusion Logic. Now ignores Light Blue background colors (common for new student rows) when building the color map.
 */

// Hardcoded sheet names (unless these are also settings, usually they are static)
const SHEET_NAMES = {
    MASTER_LIST: "Master List",
    HISTORY: "Student History"
};

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
            return "[Retention] DNC - Student explicitly opted out";
        }
    }

    // Priority 2: Waiting for Grade (Strictly for Failing Table)
    if (tableContext === 'Failing_Table' && typeof missingVal === 'number' && missingVal === 0) {
        return "[Retention] Waiting for Prof. to grade";
    }

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
 */
export async function createLDA(userOverrides, onProgress) {
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
            
            masterRange.load("values, formulas, rowIndex, columnIndex, rowCount, columnCount");
            
            const masterCellProps = masterRange.getCellProperties({
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
            
            const masterValues = masterRange.values;
            const masterFormulas = masterRange.formulas;
            const masterColors = masterCellProps.value; 
            const headers = masterValues[0];
            
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
                if (!colConfig) return -1;

                // 2. Prepare candidates list
                let aliases = [];
                if (Array.isArray(colConfig.alias)) {
                    aliases = colConfig.alias;
                } else if (colConfig.alias) {
                    aliases = [colConfig.alias];
                }

                const candidates = [colConfig.name, ...aliases];

                // 3. Find match in Excel Headers
                for (const rawCand of candidates) {
                    const candStripped = stripStr(rawCand);
                    const idx = headers.findIndex(h => stripStr(h) === candStripped);
                    if (idx !== -1) return idx;
                }

                return -1;
            };

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
            const outputColumns = settings.columns; 

            outputColumns.sort((a, b) => {
                const indexA = getColIndex(a.name);
                const indexB = getColIndex(b.name);
                if (indexA === -1 && indexB === -1) return 0;
                if (indexA === -1) return 1; 
                if (indexB === -1) return -1;
                return indexA - indexB;
            });

            const outputColIndices = outputColumns
                .map(c => getColIndex(c.name))
                .filter(idx => idx !== -1);

            // --- EXCLUDED COLORS SET ---
            // Add any other hex codes you want to ignore here.
            // Currently ignoring: White, and common Excel Light Blues
            const EXCLUDED_COLORS = new Set([
                '#ffffff', // White
                '#ADD8E6'  // Another common Google/Excel blue
            ]);

            for (let r = 1; r < masterValues.length; r++) {
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

            // --- STEP 2c: Identify Hidden Columns based on Settings ---
            const hiddenColumnsSet = new Set();
            outputColumns.forEach(c => {
                if (c.hidden) hiddenColumnsSet.add(c.name);
            });


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
                        for (let i = 1; i < hValues.length; i++) {
                            const hid = hValues[i][hIdIdx];
                            const htagRaw = String(hValues[i][hTagIdx] || '');
                            const htagLower = htagRaw.toLowerCase();
                            
                            if (hid && htagLower.includes('dnc')) {
                                dncMap.set(hid, htagLower);
                            }
                        }

                        if (settings.includeLDATag) {
                            const todayTime = new Date().setHours(0,0,0,0);
                            for (let i = hValues.length - 1; i > 0; i--) {
                                const hid = hValues[i][hIdIdx];
                                const htagRaw = String(hValues[i][hTagIdx] || '');
                                const htagLower = htagRaw.toLowerCase();
                                
                                if (hid && !ldaFollowUpMap.has(hid)) {
                                    if (htagLower.startsWith('lda ')) {
                                        const datePart = htagLower.substring(4);
                                        const ldaDate = new Date(datePart);
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

            // Determine Outreach Column Index
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
                await writeTable(
                    context, 
                    newSheet, 
                    0, 
                    "LDA_Table", 
                    outputColumns, 
                    dataRows.map(r => buildOutputRow(r, 'LDA_Table')),
                    masterSheet,
                    getColIndex,
                    dateColumnIndices
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
                
                await writeTable(
                    context, 
                    newSheet, 
                    nextRow, 
                    "Failing_Table", 
                    outputColumns, 
                    failingRows.map(r => buildOutputRow(r, 'Failing_Table')),
                    masterSheet,
                    getColIndex,
                    dateColumnIndices
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
 */
async function writeTable(context, sheet, startRow, tableName, outputColumns, processedRows, masterSheet, getColIndex, dateColumnNames) {
    if (processedRows.length === 0) return;

    const rowCount = processedRows.length;
    const colCount = outputColumns.length;
    const headers = outputColumns.map(c => c.name);

    // Prepare data blocks
    const values = [headers];
    const formulas = [headers.map(() => null)];

    processedRows.forEach(r => {
        values.push(r.cells);
        formulas.push(r.formulas);
    });

    const range = sheet.getRangeByIndexes(startRow, 0, rowCount + 1, colCount);
    range.values = values;
    
    // Create Table
    const table = sheet.tables.add(range, true);
    table.name = tableName + "_" + Math.floor(Math.random() * 1000);
    table.style = "TableStyleLight9";

    // --- FORCE Copy Conditional Formatting from Master List ---
    const cfChecks = [];
    
    outputColumns.forEach((colConfig, idx) => {
        const masterIdx = getColIndex(colConfig.name);
        if (masterIdx !== -1) {
            // Get the first data cell (Row 2, Index 1) of the master column
            const sourceCell = masterSheet.getCell(1, masterIdx);
            cfChecks.push({ 
                sourceCell, 
                targetIndex: idx,
                colName: colConfig.name
            });
        }
    });

    cfChecks.forEach(check => {
        const targetColRange = table.columns.getItemAt(check.targetIndex).getDataBodyRange();
        targetColRange.copyFrom(check.sourceCell, Excel.RangeCopyType.formats);
        
        // --- Apply Smart Date Formatting ---
        if (dateColumnNames.has(check.colName)) {
            targetColRange.numberFormat = [["mm-dd-yy;@"]];
        }
    });

    // Explicit sync to ensure formats are applied before we overlay our custom highlights
    await context.sync(); 

    // --- Apply Custom Row Colors & Cell Highlights ---
    const bodyRange = table.getDataBodyRange();
    
    processedRows.forEach((r, idx) => {
        if (r.rowColor) {
             bodyRange.getRow(idx).format.fill.color = r.rowColor;
        }
        r.cellHighlights.forEach(h => {
             const cell = bodyRange.getCell(idx, h.colIndex);
             cell.format.fill.color = h.color;
             
             if (h.strikethrough) {
                 cell.format.font.strikethrough = true;
                 cell.format.font.color = "#9C0006"; // Dark Red Text
             }
        });
        r.formulas.forEach((f, cIdx) => {
            if (f) {
                bodyRange.getCell(idx, cIdx).formulas = [[f]];
            }
        });
    });
}