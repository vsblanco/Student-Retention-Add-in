/*
 * Timestamp: 2025-11-25 17:50:00
 * Version: 2.0.0
 * Author: Gemini (for Victor)
 * Description: Core logic for creating LDA reports. Replaces old lda.js dialog logic with a module that integrates with the React UI and Workbook Settings.
 */

// Hardcoded sheet names (unless these are also settings, usually they are static)
const SHEET_NAMES = {
    MASTER_LIST: "Master List",
    HISTORY: "Student History"
};

/**
 * Main function to create the LDA report.
 * @param {Object} userOverrides - The settings from the UI (daysOut, includeFailingList, etc.)
 * @param {Function} onProgress - Callback to update UI steps: (stepId, status) => void
 */
export async function createLDA(userOverrides, onProgress) {
    try {
        // --- STEP 1: Validate Settings & Environment ---
        if (onProgress) onProgress('validate', 'active');
        
        // 1a. Load full workbook settings (including columns config)
        let workbookSettings = {};
        if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
             const settings = Office.context.document.settings.get('workbookSettings');
             if (settings && typeof settings === 'object') {
                 workbookSettings = settings;
             }
        }

        // Merge UI overrides with workbook settings defaults if needed, 
        // though usually we trust the UI for the active parameters.
        const settings = {
            daysOut: userOverrides.daysOut ?? 5,
            includeFailingList: userOverrides.includeFailingList ?? false,
            includeLDATag: userOverrides.includeLDATag ?? true,
            includeDNCTag: userOverrides.includeDNCTag ?? true,
            // Fallback to empty array if no columns configured yet
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

            // --- STEP 2: Read Master List ---
            if (onProgress) onProgress('read', 'active');

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("values, formulas, rowIndex, columnIndex, rowCount, columnCount");

            // Attempt to load history sheet for tags
            let historyData = null;
            const hasHistory = sheets.items.some(s => s.name === SHEET_NAMES.HISTORY);
            if (hasHistory) {
                const histSheet = sheets.getItem(SHEET_NAMES.HISTORY);
                const histRange = histSheet.getUsedRange();
                histRange.load("values");
                // We store the range to read it later after sync
                historyData = histRange; 
            }

            await context.sync();
            
            // Extract Headers
            const masterValues = masterRange.values;
            const masterFormulas = masterRange.formulas;
            const headers = masterValues[0]; // Raw headers from Excel
            
            // Helper to find column index based on Settings Configuration
            // Looks for a setting with id/name, then checks that setting's aliases against Excel headers
            const getColIndex = (settingName) => {
                const colConfig = settings.columns.find(c => 
                    c.name.toLowerCase() === settingName.toLowerCase()
                );
                if (!colConfig) return -1;
                
                // Aliases to check: the name itself, plus any defined aliases
                const candidates = [colConfig.name, ...(colConfig.alias || [])].map(s => String(s).trim().toLowerCase());
                
                return headers.findIndex(h => candidates.includes(String(h).trim().toLowerCase()));
            };

            // Identify Critical Columns
            const daysOutIdx = getColIndex('Days Out');
            const gradeIdx = getColIndex('Grade');
            const assignedIdx = getColIndex('Assigned');
            const studentIdIdx = getColIndex('Student Number'); // used for history matching
            const phoneIdx = getColIndex('Phone');
            const otherPhoneIdx = getColIndex('Other Phone');
            
            if (daysOutIdx === -1) throw new Error("Could not find 'Days Out' column in Master List. Check Settings.");

            if (onProgress) onProgress('read', 'completed');

            // --- STEP 3: Filtering by Days Out ---
            if (onProgress) onProgress('filter', 'active');

            // Filter rows
            const dataRows = [];
            for (let i = 1; i < masterValues.length; i++) {
                const daysOutVal = masterValues[i][daysOutIdx];
                // Check if number and meets threshold
                if (typeof daysOutVal === 'number' && daysOutVal >= settings.daysOut) {
                    dataRows.push({
                        values: masterValues[i],
                        formulas: masterFormulas[i],
                        originalIndex: i
                    });
                }
            }
            
            // Sort by Days Out descending
            dataRows.sort((a, b) => (b.values[daysOutIdx] || 0) - (a.values[daysOutIdx] || 0));
            
            if (onProgress) onProgress('filter', 'completed');

            // --- STEP 4: Filtering by Grades (Failing) ---
            // We calculate this now, but write it later
            if (onProgress) onProgress('failing', 'active');
            
            let failingRows = [];
            if (settings.includeFailingList && gradeIdx !== -1) {
                for (let i = 1; i < masterValues.length; i++) {
                    const gradeVal = masterValues[i][gradeIdx];
                    const daysOutVal = masterValues[i][daysOutIdx];
                    
                    // Logic: Grade < 60% AND Days Out <= 4 (Recent presence but failing)
                    // Handling both decimal (0.59) and whole number (59) formats just in case
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
                // Sort failing by grade ascending (lowest first)
                failingRows.sort((a, b) => (a.values[gradeIdx] || 0) - (b.values[gradeIdx] || 0));
            }
            if (onProgress) onProgress('failing', 'completed');


            // --- STEP 5: Creating Sheet ---
            if (onProgress) onProgress('createSheet', 'active');

            const today = new Date();
            const dateStr = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
            const baseName = `LDA ${dateStr}`;
            let sheetName = baseName;
            
            // Unique name logic
            let counter = 2;
            const existingNames = sheets.items.map(s => s.name);
            while (existingNames.includes(sheetName)) {
                sheetName = `${baseName} (${counter++})`;
            }

            const newSheet = sheets.add(sheetName);
            newSheet.activate();
            
            // We need to sync here to ensure sheet is created before writing
            await context.sync();
            if (onProgress) onProgress('createSheet', 'completed');


            // --- STEP 6: Applying Tags (Data Processing) ---
            if (onProgress) onProgress('tags', 'active');

            const dncMap = new Set(); // Stores IDs of DNC students
            const ldaFollowUpMap = new Map(); // Stores IDs -> { date, tag }
            
            if (historyData && studentIdIdx !== -1) {
                const hValues = historyData.values;
                if (hValues.length > 0) {
                    const hHeaders = hValues[0].map(h => String(h).toLowerCase().trim());
                    // We assume standard names for history sheet columns or scan for them
                    const hIdIdx = hHeaders.findIndex(h => h.includes('student') && h.includes('id') || h.includes('number'));
                    const hTagIdx = hHeaders.indexOf('tag');

                    if (hIdIdx !== -1 && hTagIdx !== -1) {
                        // 1. Scan for DNC
                        for (let i = 1; i < hValues.length; i++) {
                            const hid = hValues[i][hIdIdx];
                            const htag = String(hValues[i][hTagIdx] || '').toLowerCase();
                            if (hid && htag.includes('dnc')) {
                                dncMap.add(hid);
                            }
                        }

                        // 2. Scan for LDA Follow ups (Backwards for most recent)
                        if (settings.includeLDATag) {
                            const todayTime = new Date().setHours(0,0,0,0);
                            for (let i = hValues.length - 1; i > 0; i--) {
                                const hid = hValues[i][hIdIdx];
                                const htagRaw = String(hValues[i][hTagIdx] || '');
                                const htagLower = htagRaw.toLowerCase();
                                
                                if (hid && !ldaFollowUpMap.has(hid)) {
                                    // Look for "LDA <Date>"
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

            // Define output columns based on settings (filtering out hidden/static ones if needed, 
            // but usually we want to print what is in the columns setting)
            // The user settings.columns array defines the ORDER and VISIBILITY.
            const outputColumns = settings.columns.filter(c => !c.hidden);
            
            // Helper to build a table row
            const buildOutputRow = (rowObj) => {
                const cells = [];
                const formulas = [];
                let rowColor = null; // Whole row highlight (e.g. LDA follow up)
                let cellHighlights = []; // Specific cell highlights (e.g. DNC)

                const sId = rowObj.values[studentIdIdx];

                // Check LDA Follow up
                if (sId && ldaFollowUpMap.has(sId)) {
                     rowColor = "#FFEDD5"; // Light Orange
                }

                outputColumns.forEach((colConfig, colOutIdx) => {
                    // Find matching index in master data
                    const masterIdx = getColIndex(colConfig.name);
                    let val = (masterIdx !== -1) ? rowObj.values[masterIdx] : "";
                    let form = (masterIdx !== -1) ? rowObj.formulas[masterIdx] : null;

                    // Special Overrides based on Column Name
                    if (colConfig.name === 'Gradebook' && form && String(form).startsWith('=HYPERLINK')) {
                        // Keep the formula
                    } else if (colConfig.name === 'Gradebook' && val && String(val).startsWith('http')) {
                        form = `=HYPERLINK("${val}", "Link")`;
                        val = "Link";
                    }

                    // DNC Highlights
                    if (settings.includeDNCTag && dncMap.has(sId)) {
                         if (colConfig.name === 'Phone' || colConfig.name === 'Other Phone') {
                             cellHighlights.push({ colIndex: colOutIdx, color: "#FFC7CE" }); // Red for DNC
                         }
                    }

                    cells.push(val);
                    formulas.push(form);
                });

                return { cells, formulas, rowColor, cellHighlights };
            };

            // 7a. Write Main LDA Table
            if (dataRows.length > 0) {
                await writeTable(context, newSheet, 1, "LDA_Table", outputColumns.map(c => c.name), dataRows.map(buildOutputRow));
            } else {
                // Write header only if empty
                const headerRange = newSheet.getRangeByIndexes(0, 0, 1, outputColumns.length);
                headerRange.values = [outputColumns.map(c => c.name)];
            }

            // 7b. Write Failing Table (if applicable)
            let nextRow = dataRows.length + 4; // Spacing
            if (settings.includeFailingList && failingRows.length > 0) {
                const title = newSheet.getRangeByIndexes(nextRow - 1, 0, 1, 1);
                title.values = [["Failing Students (Active)"]];
                title.format.font.bold = true;
                
                await writeTable(context, newSheet, nextRow, "Failing_Table", outputColumns.map(c => c.name), failingRows.map(buildOutputRow));
            }

            // Autofit
            newSheet.getUsedRange().getEntireColumn().format.autofitColumns();

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
 * Helper to write a table and apply formatting efficiently
 */
async function writeTable(context, sheet, startRow, tableName, headers, processedRows) {
    if (processedRows.length === 0) return;

    const rowCount = processedRows.length;
    const colCount = headers.length;

    // Prepare data blocks
    const values = [headers];
    const formulas = [headers.map(() => null)]; // Header row has no formulas

    processedRows.forEach(r => {
        values.push(r.cells);
        formulas.push(r.formulas);
    });

    const range = sheet.getRangeByIndexes(startRow, 0, rowCount + 1, colCount);
    range.values = values;
    
    // Apply formulas where they exist (Excel requires non-null for formula setting usually, 
    // but mixing values/formulas in block write can be tricky. 
    // Strategy: Write values first, then iterate rows with formulas if needed, 
    // or set `formulas` property if the array is perfectly formed. 
    // For simplicity/robustness in this version, we set values. Then overlay formulas.)
    
    // Create Table
    const table = sheet.tables.add(range, true);
    table.name = tableName + "_" + Math.floor(Math.random() * 1000);
    table.style = "TableStyleLight9";

    // Batch formatting ranges
    // We can't loop every cell in Excel.run efficiently. We gather ranges.
    // However, for this size of report (usually < 200 rows), per-row formatting is acceptable if batched.
    
    // Apply Row Colors & Cell Highlights
    // We need to fetch the body range to map indexes correctly
    const bodyRange = table.getDataBodyRange();
    
    // We will apply formats in a second pass after sync if needed, 
    // but we can try to apply them using offset on the main range.
    processedRows.forEach((r, idx) => {
        if (r.rowColor) {
             bodyRange.getRow(idx).format.fill.color = r.rowColor;
        }
        r.cellHighlights.forEach(h => {
             bodyRange.getCell(idx, h.colIndex).format.fill.color = h.color;
        });
        
        // Apply formulas individually or by row to avoid overwriting values with nulls
        // (Optimization: only write if formula exists)
        r.formulas.forEach((f, cIdx) => {
            if (f) {
                bodyRange.getCell(idx, cIdx).formulas = [[f]];
            }
        });
    });
}