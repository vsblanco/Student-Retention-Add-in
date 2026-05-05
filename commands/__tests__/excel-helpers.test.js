import { describe, it, expect } from 'vitest';
import {
    findColumnIndex,
    parseHyperlinkFormula,
    normalizeHeader,
} from '../../shared/excel-helpers.js';

describe('normalizeHeader', () => {
    it('lowercases', () => {
        expect(normalizeHeader('Grade')).toBe('grade');
        expect(normalizeHeader('GRADEBOOK')).toBe('gradebook');
    });

    it('strips all whitespace (not just trim)', () => {
        expect(normalizeHeader('Grade Book')).toBe('gradebook');
        expect(normalizeHeader('  grade  book  ')).toBe('gradebook');
        expect(normalizeHeader('a\tb\nc')).toBe('abc');
    });

    it('collapses casing + whitespace differences to one canonical form', () => {
        const variants = ['Grade Book', 'gradebook', 'GRADE BOOK', '  grade book  ', 'gradeBook'];
        const normalized = variants.map(normalizeHeader);
        expect(new Set(normalized).size).toBe(1); // all collapse to 'gradebook'
    });

    it('handles null / undefined / non-string by returning empty string', () => {
        expect(normalizeHeader(null)).toBe('');
        expect(normalizeHeader(undefined)).toBe('');
        expect(normalizeHeader(42)).toBe('42'); // String() coerces, then lowercase / whitespace-strip pass through
    });

    it('NFKC-normalizes Unicode (fullwidth, ligatures)', () => {
        // Fullwidth Latin → basic Latin
        expect(normalizeHeader('Ｇｒａｄｅ Ｂｏｏｋ')).toBe('gradebook');
        // Ligature "ﬁ" → "fi"
        expect(normalizeHeader('eﬃcient')).toBe('efficient');
    });

    it('strips Unicode whitespace beyond \\s (non-breaking space, em space)', () => {
        // U+00A0 NO-BREAK SPACE between words
        expect(normalizeHeader('grade book')).toBe('gradebook');
        // U+2003 EM SPACE
        expect(normalizeHeader('a b')).toBe('ab');
    });
});

describe('findColumnIndex', () => {
    // Headers must be pre-normalized via normalizeHeader. The alias is
    // normalized inside findColumnIndex so the alias list does NOT need
    // to enumerate case or whitespace variants.
    const normalized = ['studentname', 'phone', 'email', 'lda'].map(normalizeHeader);

    it('matches an alias via normalized comparison', () => {
        expect(findColumnIndex(normalized, ['phone'])).toBe(1);
        expect(findColumnIndex(normalized, ['lda'])).toBe(3);
    });

    it('matches case variants without the alias enumerating them', () => {
        expect(findColumnIndex(normalized, ['Phone'])).toBe(1);
        expect(findColumnIndex(normalized, ['LDA'])).toBe(3);
    });

    it('matches whitespace variants without the alias enumerating them', () => {
        // Header 'studentname' (already normalized) — alias has a space
        expect(findColumnIndex(normalized, ['student name'])).toBe(0);
        expect(findColumnIndex(normalized, ['  Student Name  '])).toBe(0);
    });

    it('matches when raw headers were normalized first', () => {
        // Realistic shape: headers come from Excel with mixed case + spaces;
        // caller maps through normalizeHeader, then passes to findColumnIndex.
        const raw = ['Student Name', '  Phone  ', 'LDA'];
        const norm = raw.map(normalizeHeader);
        expect(findColumnIndex(norm, ['phone'])).toBe(1);
        expect(findColumnIndex(norm, ['student name'])).toBe(0);
    });

    it('tries aliases in order, returning the first hit', () => {
        // 'phonenumber' is not in headers; 'phone' is — should hit on 'phone'
        expect(findColumnIndex(normalized, ['phonenumber', 'phone'])).toBe(1);
    });

    it('returns -1 when no alias matches', () => {
        expect(findColumnIndex(normalized, ['ssn', 'birthday'])).toBe(-1);
    });

    it('returns -1 when headers list is empty', () => {
        expect(findColumnIndex([], ['phone'])).toBe(-1);
    });

    it('returns -1 when possibleNames is empty', () => {
        expect(findColumnIndex(normalized, [])).toBe(-1);
    });

    it('guards against non-array possibleNames input', () => {
        expect(findColumnIndex(normalized, null)).toBe(-1);
        expect(findColumnIndex(normalized, undefined)).toBe(-1);
        expect(findColumnIndex(normalized, 'phone')).toBe(-1);
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
