import ExcelJS from 'exceljs';
import { renderTemplate } from './helpers';

const HTML_TAG_RE = /<[a-z][\s\S]*?>/i;

function htmlToPlainText(value) {
    if (value === null || value === undefined) return '';
    const str = String(value);
    if (!HTML_TAG_RE.test(str)) return str;

    const doc = new DOMParser().parseFromString(`<div>${str}</div>`, 'text/html');
    doc.querySelectorAll('li').forEach(li => {
        const anchor = li.querySelector('a');
        let line;
        if (anchor) {
            const title = (anchor.textContent || '').trim();
            const url = (anchor.getAttribute('href') || '').trim();
            line = url ? `• ${title} — ${url}\n` : `• ${title}\n`;
        } else {
            line = `• ${li.textContent || ''}\n`;
        }
        li.replaceWith(doc.createTextNode(line));
    });
    doc.querySelectorAll('br').forEach(br => {
        br.replaceWith(doc.createTextNode('\n'));
    });
    return (doc.body.textContent || '').trim();
}

// StudentEmail and PersonalEmail are first-class columns so the user can pick
// whichever in Word's Send Email Messages dialog. Field name collisions (i.e. if
// the body uses {StudentEmail} or {PersonalEmail}) are deduplicated.
const FIXED_EMAIL_COLUMNS = ['StudentEmail', 'PersonalEmail'];

function uniqueFieldNames(fieldNames) {
    return fieldNames.filter(f => !FIXED_EMAIL_COLUMNS.includes(f));
}

export function buildRecipientRows(students, fieldNames) {
    const extras = uniqueFieldNames(fieldNames);
    return (students || []).map(student => {
        const row = {
            StudentEmail: student.StudentEmail || '',
            PersonalEmail: student.PersonalEmail || '',
        };
        for (const field of extras) {
            const resolved = renderTemplate(`{${field}}`, student);
            row[field] = htmlToPlainText(resolved);
        }
        return row;
    }).filter(row => row.StudentEmail);
}

export async function generateRecipientsXlsxBlob(students, fieldNames) {
    const rows = buildRecipientRows(students, fieldNames);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Recipients');

    const headers = [...FIXED_EMAIL_COLUMNS, ...uniqueFieldNames(fieldNames)];
    sheet.columns = headers.map(h => ({ header: h, key: h, width: Math.max(12, h.length + 2) }));

    for (const row of rows) {
        sheet.addRow(row);
    }

    sheet.getRow(1).font = { bold: true };
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

export async function downloadRecipientsXlsx(students, fieldNames, filename = 'email-recipients.xlsx') {
    const blob = await generateRecipientsXlsxBlob(students, fieldNames);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
