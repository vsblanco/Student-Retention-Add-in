import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import {
    buildRecipientRows,
    generateRecipientsXlsxBlob,
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

    it('strips HTML from field values and keeps URLs visible for autolinking', () => {
        const students = [
            {
                StudentEmail: 'a@x.com',
                MissingAssignmentsList: '<ul><li><a href="https://canvas/e1">Essay 1</a></li><li><a href="https://canvas/q2">Quiz 2</a></li></ul>',
            },
        ];
        const rows = buildRecipientRows(students, ['MissingAssignmentsList']);
        const text = rows[0].MissingAssignmentsList;
        expect(text).toContain('Essay 1');
        expect(text).toContain('https://canvas/e1');
        expect(text).toContain('Quiz 2');
        expect(text).toContain('https://canvas/q2');
        expect(text).not.toMatch(/<\w+/);
    });

    it('uses the unrendered placeholder when a field is missing on a student', () => {
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
