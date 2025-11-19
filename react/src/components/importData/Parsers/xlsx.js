/* * Timestamp: 2025-11-19 15:10:00 EST
 * Description: Parser for Excel (.xlsx) files using ExcelJS.
 * Returns an array of objects.
 */

import ExcelJS from 'exceljs';

export default async function parseExcel(arrayBuffer) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        
        // Use the first worksheet
        const worksheet = workbook.worksheets[0];
        if (!worksheet) return [];

        const data = [];
        let headers = [];

        // ExcelJS iterates rows (1-based index)
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const rowValues = Array.isArray(row.values) ? row.values : [];
            
            // ExcelJS row.values has an empty item at index 0 because it uses 1-based indexing
            // We slice(1) to fix this, or simply map correctly.
            
            if (rowNumber === 1) {
                // Capture Headers (trim and ensure string)
                headers = rowValues.map(v => (v == null ? '' : String(v).trim()));
            } else {
                // Capture Data
                const rowObj = {};
                // Iterate over headers to map values
                // Note: rowValues indices match headers indices roughly, but ExcelJS pads index 0
                headers.forEach((header, idx) => {
                    // header is at array index X. rowValues matches that index.
                    if (header) {
                        let val = rowValues[idx];
                        // Convert rich text/formulas to simple values if needed
                        if (val && typeof val === 'object') {
                            if (val.result !== undefined) val = val.result; // Formula result
                            else if (val.richText) val = val.richText.map(t => t.text).join(''); // Rich text
                            else if (val.text) val = val.text; // Hyperlink object
                        }
                        rowObj[header] = val == null ? '' : String(val);
                    }
                });
                // Only add if not empty
                if (Object.keys(rowObj).length > 0) {
                    data.push(rowObj);
                }
            }
        });
        
        return data;
    } catch (error) {
        console.error("Error parsing Excel file with ExcelJS:", error);
        throw error;
    }
}