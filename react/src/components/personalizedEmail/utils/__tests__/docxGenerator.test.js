import { describe, it, expect } from 'vitest';
import { Document, Paragraph } from 'docx';
import JSZip from 'jszip';
import {
    htmlToParagraphs,
    buildMailMergeTemplate,
    generateMailMergeTemplateBlob,
    extractFieldNames,
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
        expect(htmlToParagraphs('<p>First</p><p>Second</p><p>Third</p>')).toHaveLength(3);
    });

    it('produces one paragraph per <li> in a bullet list', () => {
        expect(htmlToParagraphs('<ul><li>One</li><li>Two</li><li>Three</li></ul>')).toHaveLength(3);
    });

    it('produces one paragraph per <li> in an ordered list', () => {
        expect(htmlToParagraphs('<ol><li>One</li><li>Two</li></ol>')).toHaveLength(2);
    });

    it('handles inline formatting without crashing', () => {
        const html = '<p>Hi <strong>bold</strong> and <em>italic</em> and <u>underline</u>.</p>';
        expect(htmlToParagraphs(html)).toHaveLength(1);
    });

    it('handles hyperlinks', () => {
        expect(htmlToParagraphs('<p>Visit <a href="https://example.com">our site</a> today.</p>')).toHaveLength(1);
    });

    it('handles colored text via inline style', () => {
        expect(htmlToParagraphs('<p>Notice: <span style="color: #ff0000;">red</span> here.</p>')).toHaveLength(1);
    });

    it('handles background-color highlights', () => {
        expect(htmlToParagraphs('<p><span style="background-color: rgb(255, 255, 0);">hi</span></p>')).toHaveLength(1);
    });

    it('handles {FieldName} patterns inside text', () => {
        expect(htmlToParagraphs('<p>Hi {FirstName}, you have {DaysOut} days out.</p>')).toHaveLength(1);
    });

    it('handles nested formatting', () => {
        expect(htmlToParagraphs('<p><strong><em>bold italic</em></strong></p>')).toHaveLength(1);
    });

    it('handles <br> tags inside a paragraph', () => {
        expect(htmlToParagraphs('<p>Line one<br>Line two</p>')).toHaveLength(1);
    });
});

describe('buildMailMergeTemplate', () => {
    it('returns a Document for a simple template', () => {
        expect(buildMailMergeTemplate('<p>Hi {FirstName}!</p>')).toBeInstanceOf(Document);
    });

    it('handles empty input', () => {
        expect(buildMailMergeTemplate('')).toBeInstanceOf(Document);
    });
});

describe('image handling', () => {
    it('handles <img> with a data URL', async () => {
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

    it('keeps MissingAssignmentsList as a plain MERGEFIELD (no slot expansion)', async () => {
        const blob = await generateMailMergeTemplateBlob('<p>{MissingAssignmentsList}</p>');
        const xml = await readDocumentXml(blob);
        expect(xml).toContain('MERGEFIELD MissingAssignmentsList');
    });
});
