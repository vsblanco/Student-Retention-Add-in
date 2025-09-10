import { COLUMN_MAPPINGS, SETTINGS_KEYS } from './constants.js';
import { getNameParts } from './utils.js';

/**
 * Finds the index of a column by checking against a list of possible names.
 * @param {string[]} headers An array of lowercase header names.
 * @param {string[]} possibleNames An array of possible lowercase names for the column.
 * @returns {number} The index of the column, or -1 if not found.
 */
function findColumnIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

/**
 * Fetches and processes student data from a specified worksheet.
 * @param {string} sheetName The name of the sheet to get data from.
 * @param {object[]} customParameters The array of custom parameter configurations.
 * @returns {Promise<object[]>} A promise that resolves to an array of student objects.
 */
export async function getStudentData(sheetName, customParameters) {
    console.log(`--- [LOG] Starting getStudentData for sheet: "${sheetName}" ---`);
    try {
        const values = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();
            return usedRange.values;
        });

        if (!values || values.length < 2) {
            console.warn("[LOG] No data rows found (sheet might be empty or header-only).");
            return [];
        }

        console.log("[LOG] Successfully fetched raw data from sheet.", { rows: values.length });

        const studentData = [];
        const headers = (values[0] || []).map(h => String(h || '').trim().toLowerCase());
        console.log("[LOG] Detected Headers:", headers);

        const colIndices = {};
        for (const key in COLUMN_MAPPINGS) {
            colIndices[key] = findColumnIndex(headers, COLUMN_MAPPINGS[key]);
        }
        console.log("[LOG] Column Indices Found:", colIndices);
        
        const customParamIndices = {};
        customParameters.forEach(param => {
            const headerIndex = headers.indexOf(param.sourceColumn.toLowerCase());
            if (headerIndex !== -1) {
                customParamIndices[param.name] = headerIndex;
            }
        });

        for (let i = 1; i < values.length; i++) {
            const row = values[i];
            const studentName = row[colIndices.StudentName] ?? '';
            const nameParts = getNameParts(studentName);

            const student = {
                StudentName: studentName,
                FirstName: nameParts.first,
                LastName: nameParts.last,
                StudentEmail: row[colIndices.StudentEmail] ?? '',
                PersonalEmail: row[colIndices.PersonalEmail] ?? '',
                Grade: row[colIndices.Grade] ?? '',
                DaysOut: row[colIndices.DaysOut] ?? '',
            };

            customParameters.forEach(param => {
                const colIndex = customParamIndices[param.name];
                let value = param.defaultValue ?? '';
                if (colIndex !== undefined && row[colIndex] != null) {
                    const cellValue = row[colIndex];
                    let mappingFound = false;
                    if (param.mappings && param.mappings.length > 0) {
                        for (const mapping of param.mappings) {
                            if (String(cellValue).trim().toLowerCase() === String(mapping.if).trim().toLowerCase()) {
                                value = mapping.then;
                                mappingFound = true;
                                break;
                            }
                        }
                    }
                    if (!mappingFound) {
                        value = cellValue ?? param.defaultValue ?? '';
                    }
                }
                student[param.name] = value;
            });

            studentData.push(student);
        }

        console.log(`[LOG] Finished processing. Found ${studentData.length} students.`);
        console.log(`--- [LOG] getStudentData complete ---`);
        return studentData;

    } catch (error) {
        console.error(`Error in getStudentData from sheet "${sheetName}":`, error);
        if (error.code === 'ItemNotFound') {
            throw new Error(`Sheet "${sheetName}" not found.`);
        }
        throw new Error("An error occurred while fetching student data.");
    }
}


/**
 * Builds the final JSON payload for Power Automate.
 * @param {object[]} students Array of processed student objects.
 * @param {object} template The email template object.
 * @returns {object[]} The final payload array.
 */
export function buildPayload(students, template) {
    if (!students || students.length === 0) {
        return [];
    }
    
    console.log("[LOG] Building payload for", students.length, "students.");

    const payload = students
        .map(student => ({
            from: renderTemplate(template.from, student),
            to: student.StudentEmail || '',
            cc: renderTemplate(template.cc, student),
            subject: renderTemplate(template.subject, student),
            body: renderBodyTemplate(template.body, student)
        }))
        .filter(email => email.to && email.from);

    console.log("[LOG] Final payload contains", payload.length, "valid emails.");
    return payload;
}


/**
 * Renders a simple string template with student data.
 * @param {string} templateString The string containing placeholders like {FirstName}.
 * @param {object} data The student data object.
 * @returns {string} The rendered string.
 */
function renderTemplate(templateString, data) {
    if (!templateString) return '';
    return templateString.replace(/\{(\w+)\}/g, (match, key) => {
        return (data[key] ?? match);
    });
};

/**
 * Renders the complex HTML body template, processing special tags.
 * @param {string} bodyHtml The HTML string from the Quill editor.
 * @param {object} data The student data object.
 * @returns {string} The final rendered HTML body.
 */
function renderBodyTemplate(bodyHtml, data) {
    if (!bodyHtml) return '';
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = bodyHtml;
    
    tempDiv.querySelectorAll('.parameter-tag').forEach(tag => {
        const paramName = tag.getAttribute('data-param');
        const value = data[paramName] ?? tag.innerText;
        tag.replaceWith(document.createTextNode(value));
    });

    tempDiv.querySelectorAll('.randomize-tag-wrapper').forEach(tagWrapper => {
        const options = JSON.parse(tagWrapper.dataset.options || '[]');
        const choice = options.length > 0 ? options[Math.floor(Math.random() * options.length)] : '';
        tagWrapper.replaceWith(document.createTextNode(choice));
    });

    tempDiv.querySelectorAll('.condition-tag-wrapper').forEach(tagWrapper => {
        const condition = JSON.parse(tagWrapper.dataset.condition || '{}');
        const studentValue = data[condition.if_param];
        let showText = false;
        
        const val1 = isNaN(parseFloat(studentValue)) ? String(studentValue ?? '').toLowerCase() : parseFloat(studentValue);
        const val2 = isNaN(parseFloat(condition.if_value)) ? String(condition.if_value).toLowerCase() : parseFloat(condition.if_value);

        switch(condition.operator) {
            case '=': showText = val1 == val2; break;
            case '>': showText = val1 > val2; break;
            case '>=': showText = val1 >= val2; break;
            case '<': showText = val1 < val2; break;
            case '<=': showText = val1 <= val2; break;
        }
        
        const outputText = showText ? renderTemplate(condition.then_text, data) : '';
        tagWrapper.replaceWith(document.createTextNode(outputText));
    });

    return tempDiv.innerHTML;
};


/**
 * Sends the final payload to the configured Power Automate URL.
 * @param {string} url The Power Automate HTTP endpoint.
 * @param {object[]} payload The array of email objects to send.
 */
export async function sendToPowerAutomate(url, payload) {
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }
}


// --- Workbook Settings Management ---

/**
 * Retrieves a settings object from the workbook.
 * @param {string} key The key for the setting.
 * @returns {Promise<any>} The parsed setting value, or an empty array/object.
 */
export async function getSettings(key) {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const settingItem = settings.getItemOrNullObject(key);
        settingItem.load("value");
        await context.sync();
        const defaultValue = (key === SETTINGS_KEYS.CUSTOM_PARAMS || key === SETTINGS_KEYS.EMAIL_TEMPLATES) ? [] : {};
        return settingItem.value ? JSON.parse(settingItem.value) : defaultValue;
    });
}

/**
 * Saves a settings object to the workbook.
 * @param {string} key The key for the setting.
 * @param {any} value The value to save (will be stringified).
 * @returns {Promise<any>} A promise that resolves with the saved value.
 */
export async function saveSettings(key, value) {
    await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        settings.add(key, JSON.stringify(value));
        await context.sync();
    });
    return value;
}

/**
 * Checks for a saved Power Automate connection.
 * @returns {Promise<object|null>} The connection object or null.
 */
export async function checkConnection() {
    const connections = await getSettings(SETTINGS_KEYS.CONNECTIONS);
    return (connections.length && connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email')) || null;
}

/**
 * Creates and saves a new Power Automate connection in a single Excel.run call.
 * @param {string} url The Power Automate URL.
 * @returns {Promise<object>} The new connection object.
 */
export async function createConnection(url) {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const settingItem = settings.getItemOrNullObject(SETTINGS_KEYS.CONNECTIONS);
        settingItem.load("value");
        await context.sync();

        const connections = settingItem.value ? JSON.parse(settingItem.value) : [];
        
        const newConnection = {
            id: 'pa-' + Math.random().toString(36).substr(2, 9),
            name: 'Send Personalized Email',
            type: 'power-automate',
            url: url
        };
        connections.push(newConnection);
        
        settings.add(SETTINGS_KEYS.CONNECTIONS, JSON.stringify(connections));
        await context.sync();
        
        return newConnection;
    });
}
