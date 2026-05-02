import { describe, it, expect } from 'vitest';
import {
    findColumnIndex,
    parseHyperlinkFormula,
} from '../../shared/excel-helpers.js';

describe('findColumnIndex (shared)', () => {
    const headers = ['student name', 'phone', 'email', 'lda'];

    it('lowercases the alias before matching (case-insensitive)', () => {
        expect(findColumnIndex(headers, ['Phone'])).toBe(1);
        expect(findColumnIndex(headers, ['LDA'])).toBe(3);
        expect(findColumnIndex(headers, ['Email'])).toBe(2);
    });

    it('still matches already-lowercased aliases', () => {
        expect(findColumnIndex(headers, ['phone'])).toBe(1);
    });

    it('tries aliases in order, returning the first hit', () => {
        // 'phonenumber' is not in headers; 'phone' is — should hit on 'phone'
        expect(findColumnIndex(headers, ['phonenumber', 'phone'])).toBe(1);
    });

    it('returns -1 when no alias matches', () => {
        expect(findColumnIndex(headers, ['ssn', 'birthday'])).toBe(-1);
    });

    it('returns -1 when headers list is empty', () => {
        expect(findColumnIndex([], ['phone'])).toBe(-1);
    });

    it('returns -1 when possibleNames is empty', () => {
        expect(findColumnIndex(headers, [])).toBe(-1);
    });

    it('guards against non-array possibleNames input', () => {
        expect(findColumnIndex(headers, null)).toBe(-1);
        expect(findColumnIndex(headers, undefined)).toBe(-1);
        expect(findColumnIndex(headers, 'phone')).toBe(-1);
    });
});

describe('parseHyperlinkFormula', () => {
    it('extracts url and text from a HYPERLINK formula', () => {
        expect(parseHyperlinkFormula('=HYPERLINK("https://example.com", "Click Here")'))
            .toEqual({ url: 'https://example.com', text: 'Click Here' });
    });

    it('handles no-space variant after the comma', () => {
        expect(parseHyperlinkFormula('=HYPERLINK("https://x.com","Go")'))
            .toEqual({ url: 'https://x.com', text: 'Go' });
    });

    it('is case-insensitive on the HYPERLINK keyword', () => {
        expect(parseHyperlinkFormula('=hyperlink("https://x.com", "Go")'))
            .toEqual({ url: 'https://x.com', text: 'Go' });
    });

    it('returns null for non-hyperlink strings', () => {
        expect(parseHyperlinkFormula('=SUM(A1:A10)')).toBeNull();
        expect(parseHyperlinkFormula('plain text')).toBeNull();
        expect(parseHyperlinkFormula('')).toBeNull();
    });

    it('returns null for non-string inputs', () => {
        expect(parseHyperlinkFormula(null)).toBeNull();
        expect(parseHyperlinkFormula(undefined)).toBeNull();
        expect(parseHyperlinkFormula(42)).toBeNull();
        expect(parseHyperlinkFormula({})).toBeNull();
    });

    it('returns null for HYPERLINK with only a URL (no display text)', () => {
        // Single-arg HYPERLINK (rare but possible) — current regex requires both
        expect(parseHyperlinkFormula('=HYPERLINK("https://x.com")')).toBeNull();
    });
});
