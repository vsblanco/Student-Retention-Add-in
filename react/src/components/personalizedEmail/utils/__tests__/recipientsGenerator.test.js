import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import {
    buildRecipientRows,
    generateRecipientsXlsxBlob,
    parseAssignmentList,
    computeAssignmentSlotCount,
} from '../recipientsGenerator';

function blobToArrayBuffer(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
    });
}

describe('buildRecipientRows', () => {
    it('returns one row per student with valid email', () => {
        const students = [
            { StudentEmail: 'alice@example.com', FirstName: 'Alice', Grade: 'A' },
            { StudentEmail: 'bob@example.com', FirstName: 'Bob', Grade: 'B' },
        ];
        const rows = buildRecipientRows(students, ['FirstName', 'Grade']);
        expect(rows).toEqual([
            { Email: 'alice@example.com', FirstName: 'Alice', Grade: 'A' },
            { Email: 'bob@example.com', FirstName: 'Bob', Grade: 'B' },
        ]);
    });

    it('filters out students with no email', () => {
        const students = [
            { StudentEmail: 'alice@example.com', FirstName: 'Alice' },
            { StudentEmail: '', FirstName: 'Nobody' },
            { FirstName: 'NoEmail' },
        ];
        const rows = buildRecipientRows(students, ['FirstName']);
        expect(rows).toHaveLength(1);
        expect(rows[0].Email).toBe('alice@example.com');
    });

    it('resolves nested template references in field values', () => {
        const students = [
            { StudentEmail: 'a@x.com', FirstName: 'Alice', Greeting: 'Hi {FirstName}' },
        ];
        const rows = buildRecipientRows(students, ['Greeting']);
        expect(rows[0].Greeting).toBe('Hi Alice');
    });

    it('strips HTML from field values', () => {
        const students = [
            {
                StudentEmail: 'a@x.com',
                MissingAssignmentsList: '<ul><li>Essay 1</li><li>Quiz 2</li></ul>',
            },
        ];
        const rows = buildRecipientRows(students, ['MissingAssignmentsList']);
        expect(rows[0].MissingAssignmentsList).toContain('Essay 1');
        expect(rows[0].MissingAssignmentsList).toContain('Quiz 2');
        expect(rows[0].MissingAssignmentsList).not.toMatch(/<\w+/);
    });

    it('uses empty string when a field is missing on a student', () => {
        const students = [{ StudentEmail: 'a@x.com', FirstName: 'Alice' }];
        const rows = buildRecipientRows(students, ['FirstName', 'LastName']);
        expect(rows[0].FirstName).toBe('Alice');
        expect(rows[0].LastName).toBe('{LastName}');
    });

    it('returns an empty array for null/undefined students', () => {
        expect(buildRecipientRows(null, ['x'])).toEqual([]);
        expect(buildRecipientRows(undefined, ['x'])).toEqual([]);
    });

    it('handles plain-text values without HTML processing', () => {
        const students = [{ StudentEmail: 'a@x.com', Note: 'hello world' }];
        const rows = buildRecipientRows(students, ['Note']);
        expect(rows[0].Note).toBe('hello world');
    });
});

describe('parseAssignmentList', () => {
    it('returns empty array for non-HTML input', () => {
        expect(parseAssignmentList('')).toEqual([]);
        expect(parseAssignmentList(null)).toEqual([]);
        expect(parseAssignmentList('plain text')).toEqual([]);
    });

    it('extracts title + url from anchor-wrapped list items', () => {
        const html = '<ul><li><a href="https://x.com/1">Essay 1</a></li><li><a href="https://x.com/2">Quiz 2</a></li></ul>';
        expect(parseAssignmentList(html)).toEqual([
            { title: 'Essay 1', url: 'https://x.com/1' },
            { title: 'Quiz 2', url: 'https://x.com/2' },
        ]);
    });

    it('handles list items without anchors by leaving url empty', () => {
        const html = '<ul><li>Plain Item</li><li><a href="https://x.com/1">Linked</a></li></ul>';
        expect(parseAssignmentList(html)).toEqual([
            { title: 'Plain Item', url: '' },
            { title: 'Linked', url: 'https://x.com/1' },
        ]);
    });
});

describe('computeAssignmentSlotCount', () => {
    it('returns 0 when no students have assignments', () => {
        expect(computeAssignmentSlotCount([])).toBe(0);
        expect(computeAssignmentSlotCount([{ MissingAssignmentsList: '' }])).toBe(0);
    });

    it('returns the max number of assignments across all students', () => {
        const students = [
            { MissingAssignmentsList: '<ul><li><a href="u1">a</a></li></ul>' },
            { MissingAssignmentsList: '<ul><li><a href="u1">a</a></li><li><a href="u2">b</a></li><li><a href="u3">c</a></li></ul>' },
            { MissingAssignmentsList: '<ul><li><a href="u1">a</a></li><li><a href="u2">b</a></li></ul>' },
        ];
        expect(computeAssignmentSlotCount(students)).toBe(3);
    });

    it('caps at MAX_ASSIGNMENT_SLOTS', () => {
        const items = Array.from({ length: 30 }, (_, i) => `<li><a href="u${i}">Item ${i}</a></li>`).join('');
        const students = [{ MissingAssignmentsList: `<ul>${items}</ul>` }];
        expect(computeAssignmentSlotCount(students)).toBe(20);
    });
});

describe('assignment slot expansion in rows', () => {
    it('adds A1..AN slot columns alongside a fallback MissingAssignmentsList column', async () => {
        const students = [
            {
                StudentEmail: 'a@x.com',
                FirstName: 'Alice',
                MissingAssignmentsList: '<ul><li><a href="https://canvas/e1">Essay 1</a></li><li><a href="https://canvas/q2">Quiz 2</a></li></ul>',
            },
        ];
        const blob = await generateRecipientsXlsxBlob(students, ['FirstName', 'MissingAssignmentsList'], { assignmentSlotCount: 2 });
        const buffer = await blobToArrayBuffer(blob);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const sheet = wb.getWorksheet('Recipients');
        const headers = sheet.getRow(1).values;

        expect(headers).toContain('Email');
        expect(headers).toContain('FirstName');
        expect(headers).toContain('A1_Title');
        expect(headers).toContain('A1_Url');
        expect(headers).toContain('A2_Title');
        expect(headers).toContain('A2_Url');
        // Fallback column always present — protects against placeholder-detection misses.
        expect(headers).toContain('MissingAssignmentsList');

        const row = sheet.getRow(2).values;
        const headerIdx = (name) => headers.indexOf(name);
        expect(row[headerIdx('A1_Title')]).toBe('Essay 1');
        expect(row[headerIdx('A1_Url')]).toBe('https://canvas/e1');
        expect(row[headerIdx('A2_Title')]).toBe('Quiz 2');
        expect(row[headerIdx('A2_Url')]).toBe('https://canvas/q2');
        // Fallback column includes URLs alongside titles so Outlook autolinks.
        expect(row[headerIdx('MissingAssignmentsList')]).toContain('Essay 1');
        expect(row[headerIdx('MissingAssignmentsList')]).toContain('https://canvas/e1');
    });

    it('leaves later slots empty for students with fewer assignments', async () => {
        const students = [
            {
                StudentEmail: 'a@x.com',
                MissingAssignmentsList: '<ul><li><a href="u1">One</a></li></ul>',
            },
        ];
        const blob = await generateRecipientsXlsxBlob(students, ['MissingAssignmentsList'], { assignmentSlotCount: 3 });
        const buffer = await blobToArrayBuffer(blob);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const sheet = wb.getWorksheet('Recipients');
        const headers = sheet.getRow(1).values;
        const row = sheet.getRow(2).values;
        const headerIdx = (name) => headers.indexOf(name);

        expect(row[headerIdx('A1_Title')]).toBe('One');
        expect(row[headerIdx('A2_Title')] || '').toBe('');
        expect(row[headerIdx('A3_Title')] || '').toBe('');
    });
});

describe('generateRecipientsXlsxBlob', () => {
    it('returns a Blob', async () => {
        const students = [
            { StudentEmail: 'alice@example.com', FirstName: 'Alice', Grade: 'A' },
        ];
        const blob = await generateRecipientsXlsxBlob(students, ['FirstName', 'Grade']);
        expect(blob).toBeInstanceOf(Blob);
        expect(blob.size).toBeGreaterThan(0);
    });

    it('produces a workbook with Email as the first column followed by the field columns', async () => {
        const students = [
            { StudentEmail: 'alice@example.com', FirstName: 'Alice', Grade: 'A' },
            { StudentEmail: 'bob@example.com', FirstName: 'Bob', Grade: 'B' },
        ];
        const blob = await generateRecipientsXlsxBlob(students, ['FirstName', 'Grade']);
        const buffer = await blobToArrayBuffer(blob);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const sheet = wb.getWorksheet('Recipients');
        expect(sheet).toBeDefined();

        const headerRow = sheet.getRow(1).values;
        expect(headerRow[1]).toBe('Email');
        expect(headerRow[2]).toBe('FirstName');
        expect(headerRow[3]).toBe('Grade');

        const aliceRow = sheet.getRow(2).values;
        expect(aliceRow[1]).toBe('alice@example.com');
        expect(aliceRow[2]).toBe('Alice');
        expect(aliceRow[3]).toBe('A');
    });
});
