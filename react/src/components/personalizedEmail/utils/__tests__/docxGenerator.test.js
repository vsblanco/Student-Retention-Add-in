import { describe, it, expect } from 'vitest';
import { Document, Paragraph } from 'docx';
import {
    htmlToParagraphs,
    buildEmailsDocument,
    generateEmailsDocxBlob,
} from '../docxGenerator';

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
        result.forEach(p => expect(p).toBeInstanceOf(Paragraph));
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

    it('applies firstParagraphOptions only to the first paragraph', () => {
        const result = htmlToParagraphs(
            '<p>One</p><p>Two</p>',
            { pageBreakBefore: true }
        );
        expect(result).toHaveLength(2);
        // Internal property naming on docx Paragraph is not stable across versions,
        // so just verify shape — first paragraph differs from a no-option build.
        const baseline = htmlToParagraphs('<p>One</p><p>Two</p>');
        expect(JSON.stringify(result[0])).not.toEqual(JSON.stringify(baseline[0]));
        expect(JSON.stringify(result[1])).toEqual(JSON.stringify(baseline[1]));
    });

    it('handles nested formatting', () => {
        const html = '<p><strong><em>bold italic</em></strong></p>';
        const result = htmlToParagraphs(html);
        expect(result).toHaveLength(1);
    });

    it('handles <br> tags inside a paragraph', () => {
        const html = '<p>Line one<br>Line two</p>';
        const result = htmlToParagraphs(html);
        expect(result).toHaveLength(1);
    });
});

describe('buildEmailsDocument', () => {
    it('returns a Document for a single email', () => {
        const doc = buildEmailsDocument([{ body: '<p>Hello!</p>' }]);
        expect(doc).toBeInstanceOf(Document);
    });

    it('returns a Document for multiple emails', () => {
        const doc = buildEmailsDocument([
            { body: '<p>First letter</p>' },
            { body: '<p>Second letter</p>' },
            { body: '<p>Third letter</p>' },
        ]);
        expect(doc).toBeInstanceOf(Document);
    });

    it('handles empty input', () => {
        const doc = buildEmailsDocument([]);
        expect(doc).toBeInstanceOf(Document);
    });

    it('handles emails with empty body', () => {
        const doc = buildEmailsDocument([
            { body: '' },
            { body: '<p>Real content</p>' },
        ]);
        expect(doc).toBeInstanceOf(Document);
    });
});

describe('generateEmailsDocxBlob', () => {
    it('returns a Blob for a single email', async () => {
        const blob = await generateEmailsDocxBlob([{ body: '<p>Hello world</p>' }]);
        expect(blob).toBeInstanceOf(Blob);
        expect(blob.size).toBeGreaterThan(0);
    });

    it('returns a Blob for multiple emails with mixed formatting', async () => {
        const blob = await generateEmailsDocxBlob([
            { body: '<p>Dear <strong>Alice</strong>,</p><p>Please review.</p>' },
            { body: '<p>Dear <strong>Bob</strong>,</p><ul><li>Item 1</li><li>Item 2</li></ul>' },
            { body: '<p>Dear <em>Carol</em>,</p><p>Thanks!</p>' },
        ]);
        expect(blob).toBeInstanceOf(Blob);
        expect(blob.size).toBeGreaterThan(0);
    });
});
