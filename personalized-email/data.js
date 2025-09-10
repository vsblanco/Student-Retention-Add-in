import { SETTINGS_KEYS } from './constants.js';
import { findColumnIndex, getNameParts } from './utils.js';

// --- Connection and Settings Management ---

export async function checkConnection() {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const connectionsSetting = settings.getItemOrNullObject(SETTINGS_KEYS.CONNECTIONS);
        connectionsSetting.load("value");
        await context.sync();
        const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
        return connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');
    });
}

export async function createConnection(url) {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const connectionsSetting = settings.getItemOrNullObject(SETTINGS_KEYS.CONNECTIONS);
        connectionsSetting.load("value");
        await context.sync();

        let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
        const newConnection = {
            id: 'pa-' + Math.random().toString(36).substr(2, 9),
            name: 'Send Personalized Email',
            type: 'power-automate',
            url: url,
            actions: [],
            history: []
        };
        connections.push(newConnection);
        settings.add(SETTINGS_KEYS.CONNECTIONS, JSON.stringify(connections));
        await context.sync();
        return newConnection;
    });
}

// --- Data Fetching and Processing ---

export async function getStudentData(sheetName, customParameters) {
    let studentDataCache = [];
    await Excel.run(async (context) => {
        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            const values = usedRange.values;
            if (values.length < 2) return; // No data rows

            const headers = values[0].map(h => String(h ?? '').toLowerCase());

            const colIndices = {
                StudentName: findColumnIndex(headers, ["studentname", "student name"]),
                StudentEmail: findColumnIndex(headers, ["student email", "school email", "email"]),
                PersonalEmail: findColumnIndex(headers, ["personal email", "otheremail"]),
                Grade: findColumnIndex(headers, ["grade", "course grade"]),
                DaysOut: findColumnIndex(headers, ["days out", "daysout"]),
            };

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
                    if (colIndex !== undefined) {
                        const cellValue = row[colIndex];
                        let mappingFound = false;
                        if (param.mappings && cellValue != null) {
                            for (const mapping of param.mappings) {
                                if (String(cellValue).trim().toLowerCase() === String(mapping.if).trim().toLowerCase()) {
                                    value = mapping.then;
                                    mappingFound = true;
                                    break;
                                }
                            }
                        }
                        if (!mappingFound && cellValue != null) {
                            value = cellValue ?? param.defaultValue ?? '';
                        }
                    }
                    student[param.name] = value;
                });
                studentDataCache.push(student);
            }
        } catch (error) {
            if (error.code === 'ItemNotFound') {
                throw new Error(`Sheet "${sheetName}" not found.`);
            }
            throw error;
        }
    });
    return studentDataCache;
}

export function buildPayload(students, template) {
    if (!students || students.length === 0) return [];

    return students.map(student => ({
        from: renderTemplate(template.from, student),
        to: student.StudentEmail || '',
        cc: renderTemplate(template.cc, student),
        subject: renderTemplate(template.subject, student),
        body: renderBodyTemplate(template.body, student)
    })).filter(email => email.to && email.from); // Ensure essential fields are present
}


// --- Template Rendering Logic ---

function renderTemplate(template, data) {
    if (!template) return '';
    return template.replace(/\{(\w+)\}/g, (match, key) => {
        return (data[key] ?? match);
    });
};

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

        switch (condition.operator) {
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


// --- Network ---

export async function sendToPowerAutomate(url, payload) {
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    return response;
}

// --- Template & Parameter Settings ---

export async function getSettings(key) {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const settingItem = settings.getItemOrNullObject(key);
        settingItem.load("value");
        await context.sync();
        return settingItem.value ? JSON.parse(settingItem.value) : [];
    });
}

export async function saveSettings(key, value) {
    await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        settings.add(key, JSON.stringify(value));
        await context.sync();
    });
}
