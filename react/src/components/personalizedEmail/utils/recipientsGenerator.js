import ExcelJS from 'exceljs';
import { renderTemplate } from './helpers';
import { ASSIGNMENTS_PLACEHOLDER, MAX_ASSIGNMENT_SLOTS } from './docxGenerator';

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

export function parseAssignmentList(html) {
    if (!html || typeof html !== 'string') return [];
    const doc = new DOMParser().parseFromString(`<div>${html}</div>`, 'text/html');
    const items = [];
    for (const li of doc.querySelectorAll('li')) {
        const anchor = li.querySelector('a');
        if (anchor) {
            const title = (anchor.textContent || '').trim();
            const url = anchor.getAttribute('href') || '';
            if (title || url) items.push({ title, url });
        } else {
            const title = (li.textContent || '').trim();
            if (title) items.push({ title, url: '' });
        }
    }
    return items;
}

export function computeAssignmentSlotCount(students) {
    let max = 0;
    for (const s of students || []) {
        const items = parseAssignmentList(s?.MissingAssignmentsList);
        if (items.length > max) max = items.length;
    }
    return Math.min(max, MAX_ASSIGNMENT_SLOTS);
}

function slotColumnNames(slotCount) {
    const names = [];
    for (let i = 1; i <= slotCount; i++) {
        names.push(`A${i}_Title`, `A${i}_Url`);
    }
    return names;
}

function fillAssignmentSlots(row, student, slotCount) {
    const items = parseAssignmentList(student.MissingAssignmentsList);
    for (let i = 1; i <= slotCount; i++) {
        const item = items[i - 1];
        row[`A${i}_Title`] = item?.title || '';
        row[`A${i}_Url`] = item?.url || '';
    }
}

export function buildRecipientRows(students, fieldNames, options = {}) {
    const slotCount = options.assignmentSlotCount || 0;
    // Always keep MissingAssignmentsList as a fallback column — if the docx didn't
    // expand the placeholder into slots (inline use, Quill quirks, etc.), the
    // MERGEFIELD still needs something to substitute. Slot columns are additive.
    return (students || []).map(student => {
        const row = { Email: student.StudentEmail || '' };
        for (const field of fieldNames) {
            const resolved = renderTemplate(`{${field}}`, student);
            row[field] = htmlToPlainText(resolved);
        }
        if (slotCount > 0) {
            fillAssignmentSlots(row, student, slotCount);
        }
        return row;
    }).filter(row => row.Email);
}

export async function generateRecipientsXlsxBlob(students, fieldNames, options = {}) {
    const slotCount = options.assignmentSlotCount || 0;
    const rows = buildRecipientRows(students, fieldNames, options);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Recipients');

    const headers = ['Email', ...fieldNames, ...(slotCount > 0 ? slotColumnNames(slotCount) : [])];
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

export async function downloadRecipientsXlsx(students, fieldNames, filename = 'email-recipients.xlsx', options = {}) {
    const blob = await generateRecipientsXlsxBlob(students, fieldNames, options);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
