/*
 * shared/excel-helpers.js
 *
 * Generic Excel/Office helpers shared by both runtimes.
 */

/**
 * Finds the index of a column by checking the headers array against a list
 * of possible alias names. Aliases are lowercased before matching, so
 * callers only need to pre-lowercase the headers.
 *
 * @param {string[]} headers - Lowercased header row values.
 * @param {string[]} possibleNames - Aliases to try, in order.
 * @returns {number} Matching column index, or -1 if no alias matches.
 */
export function findColumnIndex(headers, possibleNames) {
    if (!Array.isArray(possibleNames)) {
        console.error("[DEBUG] findColumnIndex received non-array for possibleNames:", possibleNames);
        return -1;
    }
    for (const name of possibleNames) {
        const index = headers.indexOf(String(name).toLowerCase());
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

/**
 * Parses an Excel HYPERLINK formula string and returns its url and display text.
 * Returns null if the string isn't a valid HYPERLINK formula.
 *
 *   parseHyperlinkFormula('=HYPERLINK("https://x.com", "Click")')
 *     → { url: 'https://x.com', text: 'Click' }
 *
 * @param {string} formula
 * @returns {{ url: string, text: string } | null}
 */
export function parseHyperlinkFormula(formula) {
    if (!formula || typeof formula !== 'string') return null;

    const hyperlinkRegex = /=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/i;
    const match = formula.match(hyperlinkRegex);

    if (match) {
        return {
            url: match[1],
            text: match[2]
        };
    }

    return null;
}
