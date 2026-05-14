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
    it('returns one row per student with both StudentEmail and PersonalEmail columns', () => {
        const students = [
            { StudentEmail: 'alice@school.edu', PersonalEmail: 'alice@home.com', FirstName: 'Alice', Grade: 'A' },
            { StudentEmail: 'bob@school.edu', PersonalEmail: '', FirstName: 'Bob', Grade: 'B' },
        ];
        const rows = buildRecipientRows(students, ['FirstName', 'Grade']);
        expect(rows).toEqual([
            { StudentEmail: 'alice@school.edu', PersonalEmail: 'alice@home.com', FirstName: 'Alice', Grade: 'A' },
            { StudentEmail: 'bob@school.edu', PersonalEmail: '', FirstName: 'Bob', Grade: 'B' },
        ]);
    });

    it('filters out students with no StudentEmail', () => {
        const students = [
            { StudentEmail: 'alice@school.edu', PersonalEmail: 'a@home.com', FirstName: 'Alice' },
            { StudentEmail: '', PersonalEmail: 'someone@home.com', FirstName: 'Nobody' },
            { FirstName: 'NoEmail' },
        ];
        const rows = buildRecipientRows(students, ['FirstName']);
        expect(rows).toHaveLength(1);
        expect(rows[0].StudentEmail).toBe('alice@school.edu');
    });

    it('does not duplicate StudentEmail/PersonalEmail when also in fieldNames', () => {
        const students = [
            { StudentEmail: 'a@x.com', PersonalEmail: 'a@home.com', FirstName: 'A' },
        ];
        const rows = buildRecipientRows(students, ['StudentEmail', 'PersonalEmail', 'FirstName']);
        expect(Object.keys(rows[0])).toEqual(['StudentEmail', 'PersonalEmail', 'FirstName']);
        expect(rows[0].StudentEmail).toBe('a@x.com');
        expect(rows[0].PersonalEmail).toBe('a@home.com');
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

    it('produces a workbook with StudentEmail and PersonalEmail columns first', async () => {
        const students = [
            { StudentEmail: 'alice@school.edu', PersonalEmail: 'alice@home.com', FirstName: 'Alice', Grade: 'A' },
            { StudentEmail: 'bob@school.edu', PersonalEmail: '', FirstName: 'Bob', Grade: 'B' },
        ];
        const blob = await generateRecipientsXlsxBlob(students, ['FirstName', 'Grade']);
        const buffer = await blobToArrayBuffer(blob);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const sheet = wb.getWorksheet('Recipients');
        expect(sheet).toBeDefined();

        const headerRow = sheet.getRow(1).values;
        expect(headerRow[1]).toBe('StudentEmail');
        expect(headerRow[2]).toBe('PersonalEmail');
        expect(headerRow[3]).toBe('FirstName');
        expect(headerRow[4]).toBe('Grade');

        const aliceRow = sheet.getRow(2).values;
        expect(aliceRow[1]).toBe('alice@school.edu');
        expect(aliceRow[2]).toBe('alice@home.com');
        expect(aliceRow[3]).toBe('Alice');
        expect(aliceRow[4]).toBe('A');

        const bobRow = sheet.getRow(3).values;
        expect(bobRow[1]).toBe('bob@school.edu');
        // PersonalEmail can be empty in ExcelJS — accept '' or undefined.
        expect(bobRow[2] || '').toBe('');
    });
});
