export default function parseCSV(text) {
	// Normalize line endings and split, skip empty trailing lines
	const lines = text.replace(/\r\n/g, '\n').split('\n').filter(l => l !== '');
	if (lines.length === 0) return [];

	const parseLine = (line) => {
		const cols = [];
		let cur = '';
		let inQuotes = false;
		for (let i = 0; i < line.length; i++) {
			const ch = line[i];
			if (ch === '"') {
				// Doubled quote -> literal quote
				if (inQuotes && line[i + 1] === '"') {
					cur += '"';
					i++;
				} else {
					inQuotes = !inQuotes;
				}
			} else if (ch === ',' && !inQuotes) {
				cols.push(cur);
				cur = '';
			} else {
				cur += ch;
			}
		}
		cols.push(cur);
		return cols;
	};

	const rows = lines.map(parseLine);
	const header = (rows.shift() || []).map(h => (h === undefined ? '' : h.trim()));
	return rows.map((row) => {
		const obj = {};
		header.forEach((h, i) => {
			obj[h || `col${i + 1}`] = row[i] !== undefined ? row[i] : '';
		});
		return obj;
	});
}
