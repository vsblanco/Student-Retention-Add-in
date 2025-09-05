export function findColumnIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
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
        // Handle "Last, First" format
        const parts = name.split(',').map(p => p.trim());
        const lastName = parts[0];
        const firstName = parts[1] || '';
        return { first: firstName, last: lastName };
    } else {
        // Handle "First Middle Last" format
        const parts = name.split(' ').filter(p => p);
        if (parts.length === 1) {
            return { first: parts[0], last: '' };
        }
        const lastName = parts.pop();
        const firstName = parts.join(' ');
        return { first: firstName, last: lastName };
    }
}

