import { describe, it, expect } from 'vitest';
import { isAuthorMatch } from '../allowlist.js';

describe('isAuthorMatch', () => {
    it('returns true when author and current user match exactly', () => {
        expect(isAuthorMatch('Jane Doe', 'Jane Doe')).toBe(true);
    });

    it('is case-insensitive', () => {
        expect(isAuthorMatch('JANE DOE', 'jane doe')).toBe(true);
        expect(isAuthorMatch('Jane Doe', 'JANE DOE')).toBe(true);
    });

    it('trims whitespace before comparing', () => {
        expect(isAuthorMatch('  Jane Doe  ', 'Jane Doe')).toBe(true);
        expect(isAuthorMatch('Jane Doe', '  Jane Doe  ')).toBe(true);
    });

    it('returns false for different names', () => {
        expect(isAuthorMatch('Jane Doe', 'John Doe')).toBe(false);
    });

    it('returns false when author is empty / null / undefined', () => {
        expect(isAuthorMatch('', 'Jane')).toBe(false);
        expect(isAuthorMatch(null, 'Jane')).toBe(false);
        expect(isAuthorMatch(undefined, 'Jane')).toBe(false);
    });

    it('returns false when current user is empty / null / undefined', () => {
        expect(isAuthorMatch('Jane', '')).toBe(false);
        expect(isAuthorMatch('Jane', null)).toBe(false);
        expect(isAuthorMatch('Jane', undefined)).toBe(false);
    });

    it('returns false when both are empty (do not consider two empties a match)', () => {
        expect(isAuthorMatch('', '')).toBe(false);
        expect(isAuthorMatch(null, null)).toBe(false);
        expect(isAuthorMatch('  ', '   ')).toBe(false);
    });

    it('coerces non-string inputs via String()', () => {
        expect(isAuthorMatch(123, '123')).toBe(true);
        expect(isAuthorMatch(0, '0')).toBe(true);
    });
});
