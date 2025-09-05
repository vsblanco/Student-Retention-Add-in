import { findColumnIndex, getTodaysLdaSheetName, getNameParts } from './utils.js';

let powerAutomateConnection = null;
let studentDataCache = [];
let lastFocusedInput = null; // To track the last focused subject/body input

const availableParameters = ['FirstName', 'LastName', 'StudentName', 'StudentEmail', 'Grade', 'DaysOut', 'Assigned'];

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("send-email-button").onclick = sendEmail;
        document.getElementById("create-connection-button").onclick = createConnection;
        document.getElementById('show-example-button').onclick = showExample;
        document.getElementById('close-example-modal-button').onclick = () => {
            document.getElementById('example-modal').classList.add('hidden');
        };

        const subjectInput = document.getElementById('email-subject');
        const bodyInput = document.getElementById('email-body');
        subjectInput.onfocus = () => lastFocusedInput = subjectInput;
        bodyInput.onfocus = () => lastFocusedInput = bodyInput;

        populateParameterButtons();
        checkConnection();
    }
});

function populateParameterButtons() {
    const container = document.getElementById('parameter-buttons');
    availableParameters.forEach(param => {
        const button = document.createElement('button');
        button.className = 'px-2 py-1 bg-gray-200 text-gray-700 text-xs rounded hover:bg-gray-300';
        button.textContent = `{${param}}`;
        button.onclick = () => insertParameter(`{${param}}`);
        container.appendChild(button);
    });
}

function insertParameter(param) {
    if (lastFocusedInput) {
        const start = lastFocusedInput.selectionStart;
        const end = lastFocusedInput.selectionEnd;
        const text = lastFocusedInput.value;
        lastFocusedInput.value = text.substring(0, start) + param + text.substring(end);
        lastFocusedInput.focus();
        lastFocusedInput.selectionStart = lastFocusedInput.selectionEnd = start + param.length;
    } else {
        // Default to body if no input has been focused yet
        const bodyInput = document.getElementById('email-body');
        bodyInput.value += param;
        bodyInput.focus();
    }
}


async function checkConnection() {
    await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const connectionsSetting = settings.getItemOrNullObject("connections");
        
        connectionsSetting.load("value");
        await context.sync();

        const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
        
        powerAutomateConnection = connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');

        if (powerAutomateConnection) {
            document.getElementById('setup-wizard').classList.add('hidden');
            document.getElementById('email-composer').classList.remove('hidden');
        } else {
            document.getElementById('setup-wizard').classList.remove('hidden');
            document.getElementById('email-composer').classList.add('hidden');
        }
    });
}

async function createConnection() {
    const urlInput = document.getElementById('power-automate-url');
    const status = document.getElementById('setup-status');
    const url = urlInput.value.trim();

    if (!isValidHttpUrl(url)) {
        status.textContent = "Please enter a valid HTTP URL.";
        status.style.color = 'red';
        return;
    }
    
    status.textContent = "Creating connection...";
    status.style.color = 'gray';

    await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const connectionsSetting = settings.getItemOrNullObject("connections");

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
        
        settings.add("connections", JSON.stringify(connections));
        await context.sync();

        status.textContent = "Connection created successfully!";
        status.style.color = 'green';

        setTimeout(checkConnection, 1500);
    });
}

async function getStudentData() {
    const recipientList = document.getElementById('recipient-list').value;
    const sheetName = recipientList === 'lda' ? getTodaysLdaSheetName() : 'Master List';
    const status = document.getElementById('status');
    status.textContent = `Fetching students from "${sheetName}"...`;
    status.style.color = 'gray';
    
    studentDataCache = []; // Clear cache before fetching

    await Excel.run(async (context) => {
        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            const values = usedRange.values;
            const headers = values[0].map(h => String(h || '').toLowerCase());
            
            const colIndices = {
                StudentName: findColumnIndex(headers, ["studentname", "student name"]),
                StudentEmail: findColumnIndex(headers, ["student email", "school email", "email"]),
                Grade: findColumnIndex(headers, ["grade", "course grade"]),
                DaysOut: findColumnIndex(headers, ["days out", "daysout"]),
                Assigned: findColumnIndex(headers, ["assigned"])
            };

            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                const studentName = row[colIndices.StudentName] || '';
                const nameParts = getNameParts(studentName);

                const student = {
                    StudentName: studentName,
                    FirstName: nameParts.first,
                    LastName: nameParts.last,
                    StudentEmail: row[colIndices.StudentEmail] || '',
                    Grade: row[colIndices.Grade] || '',
                    DaysOut: row[colIndices.DaysOut] || '',
                    Assigned: row[colIndices.Assigned] || ''
                };
                studentDataCache.push(student);
            }
            status.textContent = `Found ${studentDataCache.length} students.`;
            setTimeout(() => status.textContent = '', 3000);
        } catch (error) {
            if (error.code === 'ItemNotFound') {
                status.textContent = `Error: Sheet "${sheetName}" not found.`;
            } else {
                status.textContent = 'An error occurred while fetching data.';
            }
            status.style.color = 'red';
            console.error(error);
            throw error;
        }
    });
}

async function showExample() {
    const status = document.getElementById('status');
    try {
        await getStudentData();

        if (studentDataCache.length === 0) {
            status.textContent = 'No students found to generate an example.';
            status.style.color = 'orange';
            return;
        }

        const randomStudent = studentDataCache[Math.floor(Math.random() * studentDataCache.length)];
        
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = document.getElementById('email-body').value;

        const renderTemplate = (template, data) => {
            return template.replace(/\{(\w+)\}/g, (match, key) => {
                return data.hasOwnProperty(key) ? data[key] : match;
            });
        };

        document.getElementById('example-to').textContent = randomStudent.StudentEmail || '[No Email Found]';
        document.getElementById('example-subject').textContent = renderTemplate(subjectTemplate, randomStudent);
        document.getElementById('example-body').innerHTML = renderTemplate(bodyTemplate, randomStudent).replace(/\n/g, '<br>');

        document.getElementById('example-modal').classList.remove('hidden');

    } catch (error) {
        // Error message is already set by getStudentData
    }
}


async function sendEmail() {
    const status = document.getElementById('status');
     try {
        await getStudentData();

        if (studentDataCache.length === 0) {
            status.textContent = 'No students to send emails to.';
            status.style.color = 'orange';
            return;
        }

        status.textContent = `Preparing to send ${studentDataCache.length} emails... (Not implemented)`;
        console.log("Send Email button clicked.");
        console.log("Using Power Automate URL:", powerAutomateConnection.url);
        console.log("Recipients:", studentDataCache);

    } catch (error) {
        // Error message is already set by getStudentData
    }
}

function isValidHttpUrl(string) {
    let url;
    try {
        url = new URL(string);
    } catch (_) {
        return false;
    }
    return url.protocol === "http:" || url.protocol === "https:";
}

