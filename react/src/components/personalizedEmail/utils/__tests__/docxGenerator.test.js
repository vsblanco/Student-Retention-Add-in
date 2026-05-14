import { describe, it, expect } from 'vitest';
import { Document, Paragraph } from 'docx';
import JSZip from 'jszip';
import {
    htmlToParagraphs,
    buildMailMergeTemplate,
    generateMailMergeTemplateBlob,
    extractFieldNames,
    bodyReferencesAssignments,
} from '../docxGenerator';

function blobToArrayBuffer(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
    });
}

async function readDocumentXml(blob) {
    const buffer = await blobToArrayBuffer(blob);
    const zip = await JSZip.loadAsync(buffer);
    return zip.file('word/document.xml').async('string');
}

describe('extractFieldNames', () => {
    it('returns an empty array for empty input', () => {
        expect(extractFieldNames('')).toEqual([]);
        expect(extractFieldNames(null)).toEqual([]);
        expect(extractFieldNames(undefined)).toEqual([]);
    });

    it('extracts a single field name', () => {
        expect(extractFieldNames('Hello {FirstName}')).toEqual(['FirstName']);
    });

    it('extracts multiple distinct field names in order of first appearance', () => {
        expect(extractFieldNames('Hi {FirstName} {LastName}, your grade is {Grade}.'))
            .toEqual(['FirstName', 'LastName', 'Grade']);
    });

    it('deduplicates repeated field names but preserves first-occurrence order', () => {
        expect(extractFieldNames('{B} {A} {B} {C} {A}'))
            .toEqual(['B', 'A', 'C']);
    });

    it('ignores patterns without word characters', () => {
        expect(extractFieldNames('{} {1} {a-b}')).toEqual(['1']);
    });

    it('finds field names inside HTML', () => {
        const html = '<p>Hi <strong>{FirstName}</strong>, <em>{Grade}</em></p>';
        expect(extractFieldNames(html)).toEqual(['FirstName', 'Grade']);
    });
});

describe('htmlToParagraphs', () => {
    it('returns an empty array for empty input', () => {
        expect(htmlToParagraphs('')).toEqual([]);
        expect(htmlToParagraphs(null)).toEqual([]);
        expect(htmlToParagraphs(undefined)).toEqual([]);
    });

    it('produces a single paragraph for a plain <p>', () => {
        const result = htmlToParagraphs('<p>Hello world</p>');
        expect(result).toHaveLength(1);
        expect(result[0]).toBeInstanceOf(Paragraph);
    });

    it('produces one paragraph per <p>', () => {
        const result = htmlToParagraphs('<p>First</p><p>Second</p><p>Third</p>');
        expect(result).toHaveLength(3);
    });

    it('produces one paragraph per <li> in a bullet list', () => {
        const result = htmlToParagraphs('<ul><li>One</li><li>Two</li><li>Three</li></ul>');
        expect(result).toHaveLength(3);
    });

    it('produces one paragraph per <li> in an ordered list', () => {
        const result = htmlToParagraphs('<ol><li>One</li><li>Two</li></ol>');
        expect(result).toHaveLength(2);
    });

    it('handles inline formatting without crashing', () => {
        const html = '<p>Hi <strong>bold</strong> and <em>italic</em> and <u>underline</u>.</p>';
        const result = htmlToParagraphs(html);
        expect(result).toHaveLength(1);
    });

    it('handles hyperlinks', () => {
        const html = '<p>Visit <a href="https://example.com">our site</a> today.</p>';
        const result = htmlToParagraphs(html);
        expect(result).toHaveLength(1);
    });

    it('handles colored text via inline style', () => {
        const html = '<p>Notice: <span style="color: #ff0000;">red text</span> here.</p>';
        const result = htmlToParagraphs(html);
        expect(result).toHaveLength(1);
    });

    it('handles background-color highlights', () => {
        const html = '<p>This is <span style="background-color: rgb(255, 255, 0);">highlighted</span>.</p>';
        const result = htmlToParagraphs(html);
        expect(result).toHaveLength(1);
    });

    it('handles {FieldName} patterns inside text', () => {
        const result = htmlToParagraphs('<p>Hi {FirstName}, you have {DaysOut} days out.</p>');
        expect(result).toHaveLength(1);
    });

    it('handles nested formatting', () => {
        const result = htmlToParagraphs('<p><strong><em>bold italic</em></strong></p>');
        expect(result).toHaveLength(1);
    });

    it('handles <br> tags inside a paragraph', () => {
        const result = htmlToParagraphs('<p>Line one<br>Line two</p>');
        expect(result).toHaveLength(1);
    });
});

describe('buildMailMergeTemplate', () => {
    it('returns a Document for a simple template', () => {
        const doc = buildMailMergeTemplate('<p>Hi {FirstName}!</p>');
        expect(doc).toBeInstanceOf(Document);
    });

    it('handles empty input', () => {
        const doc = buildMailMergeTemplate('');
        expect(doc).toBeInstanceOf(Document);
    });

    it('handles templates with no merge fields', () => {
        const doc = buildMailMergeTemplate('<p>Static greeting, no personalization.</p>');
        expect(doc).toBeInstanceOf(Document);
    });
});

describe('bodyReferencesAssignments', () => {
    it('returns true when the body contains the MissingAssignmentsList placeholder', () => {
        expect(bodyReferencesAssignments('<p>Hi {MissingAssignmentsList}</p>')).toBe(true);
    });
    it('returns false otherwise', () => {
        expect(bodyReferencesAssignments('<p>Hi {FirstName}</p>')).toBe(false);
        expect(bodyReferencesAssignments('')).toBe(false);
        expect(bodyReferencesAssignments(null)).toBe(false);
    });
});

describe('assignment slot expansion', () => {
    it('produces N paragraphs for N slots when body has placeholder on its own line', async () => {
        const html = '<p>Hi {FirstName}!</p><p>{MissingAssignmentsList}</p><p>Bye!</p>';
        const blob = await generateMailMergeTemplateBlob(html, { assignmentSlotCount: 5 });
        const xml = await readDocumentXml(blob);
        const ifCount = (xml.match(/MERGEFIELD A\d+_Title/g) || []).length;
        expect(ifCount).toBeGreaterThanOrEqual(10);
    });

    it('emits HYPERLINK + MERGEFIELD URL pairs', async () => {
        const html = '<p>{MissingAssignmentsList}</p>';
        const blob = await generateMailMergeTemplateBlob(html, { assignmentSlotCount: 2 });
        const xml = await readDocumentXml(blob);
        expect(xml).toContain('HYPERLINK');
        expect(xml).toContain('MERGEFIELD A1_Url');
        expect(xml).toContain('MERGEFIELD A2_Url');
        expect(xml).toContain('MERGEFIELD A1_Title');
        expect(xml).toContain('MERGEFIELD A2_Title');
    });

    it('emits IF guards so empty slots collapse', async () => {
        const html = '<p>{MissingAssignmentsList}</p>';
        const blob = await generateMailMergeTemplateBlob(html, { assignmentSlotCount: 3 });
        const xml = await readDocumentXml(blob);
        const ifCount = (xml.match(/\sIF\s/g) || []).length;
        expect(ifCount).toBeGreaterThanOrEqual(3);
    });

    it('does not expand when assignmentSlotCount is 0', async () => {
        const html = '<p>{MissingAssignmentsList}</p>';
        const blob = await generateMailMergeTemplateBlob(html, { assignmentSlotCount: 0 });
        const xml = await readDocumentXml(blob);
        expect(xml).toContain('MERGEFIELD MissingAssignmentsList');
        expect(xml).not.toContain('A1_Url');
    });

    it('expands placeholder even when Quill adds a trailing zero-width space', async () => {
        // Reproduces the real-world Quill output: parameter wrapped in a span,
        // followed by a zero-width space "buffer" character that Quill inserts
        // so the cursor doesn't pick up the highlight style.
        const html = '<p><span>{MissingAssignmentsList}</span>​</p>';
        const blob = await generateMailMergeTemplateBlob(html, { assignmentSlotCount: 2 });
        const xml = await readDocumentXml(blob);
        expect(xml).toContain('MERGEFIELD A1_Title');
        expect(xml).toContain('MERGEFIELD A2_Title');
    });

    it('expands placeholder when wrapped in a styled span (Quill default)', async () => {
        const html = '<p><span style="background-color: rgb(254, 215, 170);">{MissingAssignmentsList}</span></p>';
        const blob = await generateMailMergeTemplateBlob(html, { assignmentSlotCount: 3 });
        const xml = await readDocumentXml(blob);
        expect(xml).toContain('MERGEFIELD A1_Title');
    });
});

describe('image handling', () => {
    it('handles <img> with a data URL', async () => {
        // 1x1 transparent PNG
        const onePxPng = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
        const html = `<p>Logo: <img src="data:image/png;base64,${onePxPng}" width="50" height="50"/></p>`;
        const blob = await generateMailMergeTemplateBlob(html);
        const buffer = await blobToArrayBuffer(blob);
        const zip = await JSZip.loadAsync(buffer);
        const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('word/media/'));
        expect(mediaFiles.length).toBeGreaterThan(0);
    });

    it('skips <img> tags whose data could not be fetched', async () => {
        const html = '<p>Logo: <img src="https://example.invalid/missing.png" width="50" height="50"/></p>';
        const blob = await generateMailMergeTemplateBlob(html);
        expect(blob.size).toBeGreaterThan(0);
    });
});

describe('generateMailMergeTemplateBlob', () => {
    it('returns a Blob for a template with merge fields', async () => {
        const blob = await generateMailMergeTemplateBlob(
            '<p>Dear {FirstName},</p><p>Your grade is <strong>{Grade}</strong>.</p>'
        );
        expect(blob).toBeInstanceOf(Blob);
        expect(blob.size).toBeGreaterThan(0);
    });

    it('emits MERGEFIELD instruction text in the docx XML', async () => {
        const blob = await generateMailMergeTemplateBlob('<p>Hi {FirstName}, your grade is {Grade}.</p>');
        const xml = await readDocumentXml(blob);
        expect(xml).toContain('MERGEFIELD FirstName');
        expect(xml).toContain('MERGEFIELD Grade');
    });

    it('emits one MERGEFIELD per occurrence including repeats', async () => {
        const blob = await generateMailMergeTemplateBlob('<p>{FirstName} {FirstName} {LastName}</p>');
        const xml = await readDocumentXml(blob);
        const firstNameCount = (xml.match(/MERGEFIELD FirstName/g) || []).length;
        const lastNameCount = (xml.match(/MERGEFIELD LastName/g) || []).length;
        expect(firstNameCount).toBe(2);
        expect(lastNameCount).toBe(1);
    });
});
