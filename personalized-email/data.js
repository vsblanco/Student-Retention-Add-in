import { SETTINGS_KEYS, COLUMN_MAPPINGS } from './constants.js';
import { getNameParts, findColumnIndex } from './utils.js';

// --- Connection and Settings ---

export async function checkConnection() {
    const connections = await getSettings(SETTINGS_KEYS.CONNECTIONS);
    return connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');
}

export async function createConnection(url) {
    const connections = await getSettings(SETTINGS_KEYS.CONNECTIONS);
    const newConnection = {
        id: 'pa-' + Math.random().toString(36).substr(2, 9),
        name: 'Send Personalized Email',
        type: 'power-automate',
        url: url,
        actions: [],
        history: []
    };
    connections.push(newConnection);
    await saveSettings(SETTINGS_KEYS.CONNECTIONS, connections);
    return newConnection;
}

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

// --- Data Fetching and Processing ---

export async function getStudentData(sheetName, customParameters) {
    const studentData = [];
    await Excel.run(async (context) => {
        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            const values = usedRange.values;
            if (values.length < 2) {
                // Not enough data to process (must have header + at least one student)
                return;
            }

            const headers = values[0].map(h => String(h ?? '').toLowerCase());
            
            // Use the flexible COLUMN_MAPPINGS from constants to find indices
            const colIndices = {
                StudentName: findColumnIndex(headers, COLUMN_MAPPINGS.StudentName),
                StudentEmail: findColumnIndex(headers, COLUMN_MAPPINGS.StudentEmail),
                PersonalEmail: findColumnIndex(headers, COLUMN_MAPPINGS.PersonalEmail),
                Grade: findColumnIndex(headers, COLUMN_MAPPINGS.Grade),
                DaysOut: findColumnIndex(headers, COLUMN_MAPPINGS.DaysOut),
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
                studentData.push(student);
            }
        } catch (error) {
            if (error.code === 'ItemNotFound') {
                throw new Error(`Sheet "${sheetName}" not found.`);
            } else {
                console.error(error);
                throw new Error('An error occurred while fetching student data.');
            }
        }
    });
    return studentData;
}

export function buildPayload(students, template) {
    if (!students || students.length === 0) {
        return [];
    }
    
    const renderTemplate = (templateStr, data) => {
        if (!templateStr) return '';
        return templateStr.replace(/\{(\w+)\}/g, (match, key) => data[key] ?? match);
    };

    const renderBodyTemplate = (bodyHtml, data) => {
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
            
            if (studentValue !== undefined && studentValue !== null) {
                const val1 = isNaN(parseFloat(studentValue)) ? String(studentValue).trim().toLowerCase() : parseFloat(studentValue);
                const val2 = isNaN(parseFloat(condition.if_value)) ? String(condition.if_value).trim().toLowerCase() : parseFloat(condition.if_value);

                switch(condition.operator) {
                    case '=': showText = val1 == val2; break;
                    case '>': showText = val1 > val2; break;
                    case '>=': showText = val1 >= val2; break;
                    case '<': showText = val1 < val2; break;
                    case '<=': showText = val1 <= val2; break;
                }
            }
            
            const outputText = showText ? renderTemplate(condition.then_text, data) : '';
            tagWrapper.replaceWith(document.createTextNode(outputText));
        });
        return tempDiv.innerHTML;
    };
    
    return students.map(student => ({
        from: renderTemplate(template.from, student),
        to: student.StudentEmail || '',
        cc: renderTemplate(template.cc, student),
        subject: renderTemplate(template.subject, student),
        body: renderBodyTemplate(template.body, student)
    })).filter(email => email.to && email.from); // Filter out emails with no recipient or sender
}

export async function sendToPowerAutomate(url, payload) {
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    return response;
}

