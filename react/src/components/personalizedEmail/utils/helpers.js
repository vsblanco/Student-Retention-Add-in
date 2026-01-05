// Utility functions for personalized email feature

export function findColumnIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const index = headers.indexOf(name.toLowerCase());
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

export function getTodaysLdaSheetName() {
    const now = new Date();
    return `LDA ${now.getMonth() + 1}-${now.getDate()}-${now.getFullYear()}`;
}

export function getNameParts(fullName) {
    if (!fullName || typeof fullName !== 'string') {
        return { first: '', last: '' };
    }

    const name = fullName.trim();

    if (name.includes(',')) {
        const parts = name.split(',').map(p => p.trim());
        const lastName = parts[0];
        const firstName = parts[1] || '';
        return { first: firstName, last: lastName };
    } else {
        const parts = name.split(' ').filter(p => p);
        if (parts.length === 1) {
            return { first: parts[0], last: '' };
        }
        const lastName = parts.pop();
        const firstName = parts.join(' ');
        return { first: firstName, last: lastName };
    }
}

export function isValidEmail(email) {
    if (typeof email !== 'string' || !email.trim()) {
        return false;
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

export function isValidHttpUrl(string) {
    try {
        const url = new URL(string);
        return url.protocol === "http:" || url.protocol === "https:";
    } catch (_) {
        return false;
    }
}

export function evaluateMapping(cellValue, mapping) {
    const cellStr = String(cellValue).trim().toLowerCase();
    const conditionStr = String(mapping.if).trim().toLowerCase();
    const cellNum = parseFloat(cellValue);
    const conditionNum = parseFloat(mapping.if);
    const isNumeric = !isNaN(cellNum) && !isNaN(conditionNum);

    switch (mapping.operator) {
        case 'eq': return cellStr === conditionStr;
        case 'neq': return cellStr !== conditionStr;
        case 'contains': return cellStr.includes(conditionStr);
        case 'does_not_contain': return !cellStr.includes(conditionStr);
        case 'starts_with': return cellStr.startsWith(conditionStr);
        case 'ends_with': return cellStr.endsWith(conditionStr);
        case 'gt': return isNumeric && cellNum > conditionNum;
        case 'lt': return isNumeric && cellNum < conditionNum;
        case 'gte': return isNumeric && cellNum >= conditionNum;
        case 'lte': return isNumeric && cellNum <= conditionNum;
        default: return false;
    }
}

export const renderTemplate = (template, data) => {
    if (!template) return '';
    let result = template;
    for (let i = 0; i < 10 && /\{(\w+)\}/.test(result); i++) {
        result = result.replace(/\{(\w+)\}/g, (match, key) => {
            let valueToInsert = data.hasOwnProperty(key) ? data[key] : match;
            if (typeof valueToInsert === 'string') {
                const trimmedValue = valueToInsert.trim();
                if (trimmedValue.startsWith('<p>') && trimmedValue.endsWith('</p>')) {
                    const innerHtml = trimmedValue.substring(3, trimmedValue.length - 4);
                    if (!innerHtml.includes('<p>') && !innerHtml.includes('<div>')) {
                        valueToInsert = innerHtml;
                    }
                }
            }
            return valueToInsert;
        });
    }
    return result;
};

export const renderCCTemplate = (recipients, data) => {
    if (!recipients || recipients.length === 0) return '';
    return recipients.map(recipient => renderTemplate(recipient, data)).join(';');
};

/**
 * Extracts the URL and display text from an Excel HYPERLINK formula
 * Example: =HYPERLINK("https://example.com", "Click Here") => { url: "https://example.com", text: "Click Here" }
 */
export function parseHyperlinkFormula(formula) {
    if (!formula || typeof formula !== 'string') return null;

    // Match HYPERLINK formula pattern: =HYPERLINK("url", "display text")
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

/**
 * Generates an HTML list of missing assignments for a student
 * Returns empty string if no assignments found
 */
export async function generateMissingAssignmentsList(gradeBookValue, gradeBookFormula, context) {
    try {
        // Extract the Grade Book identifier from the formula or value
        let gradeBookId = gradeBookValue;
        if (gradeBookFormula) {
            const parsed = parseHyperlinkFormula(gradeBookFormula);
            if (parsed) {
                gradeBookId = parsed.text; // Use the display text as identifier
            }
        }

        if (!gradeBookId) return '';

        // Access the Missing Assignments sheet
        const missingSheet = context.workbook.worksheets.getItem("Missing Assignments");
        const usedRange = missingSheet.getUsedRange();
        usedRange.load("values, formulas");
        await context.sync();

        const values = usedRange.values;
        const formulas = usedRange.formulas;
        const headers = values[0].map(h => String(h ?? '').toLowerCase());

        // Find column indices
        const gradeBookColIndex = headers.findIndex(h =>
            h.includes('grade book') || h.includes('gradebook')
        );
        const assignmentColIndex = headers.findIndex(h =>
            h.includes('assignment')
        );

        if (gradeBookColIndex === -1 || assignmentColIndex === -1) {
            return '';
        }

        // Find all matching assignments
        const assignments = [];
        for (let i = 1; i < values.length; i++) {
            const rowGradeBookValue = values[i][gradeBookColIndex];
            const rowGradeBookFormula = formulas[i][gradeBookColIndex];

            // Extract the display text if it's a HYPERLINK formula
            let rowGradeBookId = rowGradeBookValue;
            if (rowGradeBookFormula) {
                const parsed = parseHyperlinkFormula(rowGradeBookFormula);
                if (parsed) {
                    rowGradeBookId = parsed.text; // Use display text for comparison
                }
            }

            // Check if this row matches the student's grade book
            if (String(rowGradeBookId).trim() === String(gradeBookId).trim()) {
                const assignmentFormula = formulas[i][assignmentColIndex];
                const assignmentValue = values[i][assignmentColIndex];

                // Parse the assignment hyperlink
                const parsed = parseHyperlinkFormula(assignmentFormula);
                if (parsed) {
                    assignments.push({
                        url: parsed.url,
                        title: parsed.text
                    });
                } else if (assignmentValue) {
                    // Fallback if no hyperlink
                    assignments.push({
                        url: null,
                        title: String(assignmentValue)
                    });
                }
            }
        }

        if (assignments.length === 0) {
            return '';
        }

        // Generate HTML bullet list
        const listItems = assignments.map(assignment => {
            if (assignment.url) {
                return `<li><a href="${assignment.url}" target="_blank">${assignment.title}</a></li>`;
            } else {
                return `<li>${assignment.title}</li>`;
            }
        }).join('');

        return `<ul>${listItems}</ul>`;

    } catch (error) {
        console.error('Error generating missing assignments list:', error);
        return '';
    }
}
