// PERF: small in-memory cache for normalized keys to avoid repeated work
const _normCache = new Map();

export function normalizeKey(v) {
	if (v == null) return '';
	// use trimmed string as cache key to avoid repeated trim/lower/replace
	const raw = String(v).trim();
	const cached = _normCache.get(raw);
	if (cached !== undefined) return cached;
	const norm = raw.toLowerCase().replace(/\s+/g, '');
	_normCache.set(raw, norm);
	return norm;
}

// NEW helper: build fast lookup map of normalized header -> index
export function getHeaderIndexMap(headerArr = []) {
	const map = new Map();
	for (let i = 0; i < (headerArr || []).length; i++) {
		const v = headerArr[i];
		// trim once and normalizeKey uses cache
		map.set(normalizeKey(v == null ? '' : String(v).trim()), i);
	}
	return map;
}

// NEW helper: extract Hyperlink from Excel HYPERLINK formula or plain URL
export function parseHyperlink(formulaOrValue) {
	// returns { url: string|null, display: string|null }
	if (!formulaOrValue) return { url: null, display: null };
	if (typeof formulaOrValue !== 'string') return { url: null, display: null };
	const s = formulaOrValue.trim();
	// match HYPERLINK("url","display") where either quote can be ' or "
	const m = s.match(/=\s*HYPERLINK\s*\(\s*["']([^"']+)["']\s*,\s*["']([^"']*)["']\s*\)/i);
	if (m) return { url: m[1].trim(), display: m[2].trim() };
	// fallback: plain URL value -> use as both url and display
	if (/^https?:\/\//i.test(s)) return { url: s, display: s };
	return { url: null, display: null };
}

export function extractHyperlink(formulaOrValue) {
	const p = parseHyperlink(formulaOrValue);
	return p && p.url ? p.url : null;
}

export function makeHyperlinkFormula(url, display) {
	// Escape internal double-quotes for Excel by doubling them
	const escUrl = String(url).replace(/"/g, '""');
	const escDisplay = (display == null ? escUrl : String(display).replace(/"/g, '""'));
	return `=HYPERLINK("${escUrl}","${escDisplay}")`;
}

// NEW helper: find column index in a header array for a settings entry (match name first, then aliases)
export function findColumnIndex(headerArr = [], setting = {}) {
	if (!Array.isArray(headerArr) || headerArr.length === 0 || !setting) return -1;
	const indexMap = getHeaderIndexMap(headerArr);

	// check canonical name first
	if (setting.name) {
		const nk = normalizeKey(setting.name);
		if (indexMap.has(nk)) return indexMap.get(nk);
	}
	// then check aliases
	if (setting.alias) {
		if (Array.isArray(setting.alias)) {
			for (const a of setting.alias) {
				const nk = normalizeKey(a);
				if (indexMap.has(nk)) return indexMap.get(nk);
			}
		} else {
			const nk = normalizeKey(setting.alias);
			if (indexMap.has(nk)) return indexMap.get(nk);
		}
	}
	return -1;
}

// buildAliasMap: returns a Map where normalized target column -> canonical setting.name
export function buildAliasMap(settingsCols = [], targetCols = []) {
	const targetSet = new Set((targetCols || []).map(normalizeKey));
	const map = new Map();

	(settingsCols || []).forEach((sc) => {
		if (!sc) return;
		const canonical = sc.name ? String(sc.name).trim() : null;
		if (!canonical) return;
		const candidates = new Set();
		candidates.add(normalizeKey(canonical));
		if (sc.alias) {
			if (Array.isArray(sc.alias)) sc.alias.forEach((a) => candidates.add(normalizeKey(a)));
			else candidates.add(normalizeKey(sc.alias));
		}
		// find which target column matches any candidate, map that target to canonical
		for (const cand of candidates) {
			if (targetSet.has(cand)) {
				// map the target (normalized) -> canonical (original casing from settings)
				map.set(cand, canonical);
			}
		}
	});

	return map;
}

// rename header array using aliasMap (map keys are normalized target -> canonical)
export function renameHeaderArray(headerArr = [], aliasMap = new Map()) {
	return (headerArr || []).map((h) => {
		const key = normalizeKey(h);
		return aliasMap.has(key) ? aliasMap.get(key) : h;
	});
}

// gatherIdentifierColumn: find settingsColumns entry where identifier === true
export function gatherIdentifierColumn(settingsCols = []) {
	const cols = Array.isArray(settingsCols) ? settingsCols : [];
	if (!cols || cols.length === 0) {
		console.log('gatherIdentifierColumn: no settings columns provided');
		return null;
	}
	// support boolean true, string "true", and common misspelling 'identifer'
	const found = cols.find((c) => {
		if (!c || typeof c !== 'object') return false;
		const id = c.identifier;
		const idMiss = c.identifer; // tolerate misspelling seen elsewhere
		return id === true || id === 'true' || idMiss === true || idMiss === 'true';
	}) || null;
	console.log('gatherIdentifierColumn found:', found);
	return found;
}

// rename object keys for array of objects using aliasMap
export function renameObjectRows(rows = [], aliasMap = new Map()) {
	if (!aliasMap || aliasMap.size === 0) return rows.slice();
	return rows.map((obj) => {
		if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return obj;
		const newObj = {};
		Object.keys(obj).forEach((k) => {
			const nk = normalizeKey(k);
			const canonical = aliasMap.get(nk);
			if (canonical) newObj[canonical] = obj[k];
			else newObj[k] = obj[k];
		});
		return newObj;
	});
}

// rename first-row header of array-of-arrays and keep remaining rows unchanged
export function renameArrayRows(rows = [], aliasMap = new Map()) {
	if (!Array.isArray(rows) || rows.length === 0 || aliasMap.size === 0) return rows.slice();
	const header = rows[0];
	const renamedHeader = renameHeaderArray(header, aliasMap);
	return [renamedHeader, ...rows.slice(1)];
}

// NEW helper: compute saved static data from a used-range values array (no Excel.run)
export function computeSavedStaticFromValues(values = [], sheetHeaders = [], staticCols = [], identifierSetting = null, formulas = []) {
	if (!Array.isArray(staticCols) || staticCols.length === 0) return null;
	if (!identifierSetting || !identifierSetting.name) return null;
	if (!Array.isArray(values) || values.length === 0) return null;

	const headerRow = (Array.isArray(sheetHeaders) && sheetHeaders.length > 0)
		? sheetHeaders.map((v) => (v == null ? '' : String(v).trim()))
		: (Array.isArray(values[0]) ? values[0].map((v) => (v == null ? '' : String(v).trim())) : []);

	const headerIndexMap = getHeaderIndexMap(headerRow);

	const identifierIndex = headerIndexMap.has(normalizeKey(identifierSetting.name))
		? headerIndexMap.get(normalizeKey(identifierSetting.name))
		: -1;
	if (identifierIndex === -1) return null;

	const staticIndices = staticCols.map((colName) => (headerIndexMap.has(normalizeKey(colName)) ? headerIndexMap.get(normalizeKey(colName)) : -1));

	const saved = new Map();
	const formulasIsArray = Array.isArray(formulas);

	for (let r = 1; r < values.length; r++) {
		const row = Array.isArray(values[r]) ? values[r] : [];
		let idValRaw = row[identifierIndex];
		let idVal = idValRaw == null ? '' : String(idValRaw).trim();

		if (formulasIsArray && formulas[r] && typeof formulas[r][identifierIndex] === 'string' && formulas[r][identifierIndex].trim().startsWith('=')) {
			const parsedId = parseHyperlink(formulas[r][identifierIndex]);
			if (parsedId.url) idVal = parsedId.url;
		}

		const obj = {};

		for (let i = 0; i < staticCols.length; i++) {
			const colName = staticCols[i];
			const colIdx = staticIndices[i];
			if (colIdx < 0) continue;
			let rawVal = colIdx >= 0 && row.length > colIdx ? row[colIdx] : null;

			// if the cell had a formula with HYPERLINK, recreate the full HYPERLINK(...) formula
			if (colIdx >= 0 && formulasIsArray && formulas[r] && typeof formulas[r][colIdx] === 'string' && formulas[r][colIdx].trim().startsWith('=')) {
				const parsed = parseHyperlink(formulas[r][colIdx]);
				if (parsed.url) {
					rawVal = makeHyperlinkFormula(parsed.url, parsed.display);
				}
			}

			if (rawVal != null) {
				const asStr = String(rawVal).trim();
				if (asStr !== '') {
					obj[colName] = rawVal;
				}
			}
		}

		if (Object.keys(obj).length > 0) {
			saved.set(String(idVal), obj);
		}
	}

	// LOG: report computeSavedStaticFromValues summary
	if (saved.size > 0) {
		console.log(`computeSavedStaticFromValues: captured ${saved.size} static row entries for cols: ${staticCols.join(', ')}`);
	} else {
		console.log('computeSavedStaticFromValues: no static values captured');
	}

	return {
		identifierIndex,
		staticCols,
		staticIndices,
		savedMap: saved,
	};
}

// NEW helper: compute identifier index and Set of identifiers from a used-range values array (no Excel.run)
export function computeIdentifierListFromValues(values = [], sheetHeaders = [], identifierSetting = null, formulas = []) {
	if (!identifierSetting || !identifierSetting.name) return { identifierIndex: -1, identifiers: new Set() };
	if (!Array.isArray(values) || values.length === 0) return { identifierIndex: -1, identifiers: new Set() };

	const headerRow = (Array.isArray(sheetHeaders) && sheetHeaders.length > 0)
		? sheetHeaders.map((v) => (v == null ? '' : String(v).trim()))
		: (Array.isArray(values[0]) ? values[0].map((v) => (v == null ? '' : String(v).trim())) : []);

	const headerIndexMap = getHeaderIndexMap(headerRow);

	const identifierIndex = headerIndexMap.has(normalizeKey(identifierSetting.name))
		? headerIndexMap.get(normalizeKey(identifierSetting.name))
		: -1;
	if (identifierIndex === -1) return { identifierIndex: -1, identifiers: new Set() };

	const ids = new Set();
	const formulasIsArray = Array.isArray(formulas);

	for (let r = 1; r < values.length; r++) {
		const row = Array.isArray(values[r]) ? values[r] : [];
		let raw = row.length > identifierIndex ? row[identifierIndex] : null;
		if (formulasIsArray && formulas[r] && typeof formulas[r][identifierIndex] === 'string' && formulas[r][identifierIndex].trim().startsWith('=')) {
			const parsed = parseHyperlink(formulas[r][identifierIndex]);
			if (parsed.url) raw = parsed.url;
		}
		const id = raw == null ? '' : String(raw).trim();
		if (id !== '') ids.add(String(id));
	}

	return { identifierIndex, identifiers: ids };
}

// NEW helper: apply static columns inside an existing Excel.run context and sheet
// Accepts optional sheetName for logging clarity.
export async function applyStaticColumnsWithContext(context, sheet, savedInfo, writeStartRow = 0, rowCount = 0, sheetName = '') {
	if (!savedInfo || !savedInfo.savedMap || savedInfo.savedMap.size === 0) return;
	if (!sheet || rowCount <= 0) return;

	// read identifiers for the newly written rows
	const idRange = sheet.getRangeByIndexes(writeStartRow, savedInfo.identifierIndex, rowCount, 1);
	idRange.load('values');
	await context.sync();

	const idValues = idRange.values || [];
	const savedMap = savedInfo.savedMap;
	const staticCols = savedInfo.staticCols || [];
	const staticIndices = savedInfo.staticIndices || [];

	for (let i = 0; i < staticCols.length; i++) {
		const colName = staticCols[i];
		const colIndex = staticIndices[i];
		if (colIndex === -1) continue;

		// Build column vector of values and formula flags
		const valuesToWrite = new Array(rowCount);
		const isFormula = new Array(rowCount);
		for (let r = 0; r < rowCount; r++) {
			const rawId = idValues[r] && idValues[r][0] != null ? String(idValues[r][0]).trim() : '';
			const savedRow = savedMap.get(String(rawId));
			const v = savedRow && Object.prototype.hasOwnProperty.call(savedRow, colName) ? savedRow[colName] : '';
			valuesToWrite[r] = v;
			isFormula[r] = (typeof v === 'string' && v.startsWith('='));
		}

		// Coalesce contiguous runs of same type and write each run once
		let runStart = 0;
		let runType = isFormula[0] || false; // false => values, true => formulas
		for (let r = 0; r <= rowCount; r++) {
			const curType = r < rowCount ? isFormula[r] : null;
			if (r === rowCount || curType !== runType) {
				const runLen = r - runStart;
				if (runLen > 0) {
					const startAbs = writeStartRow + runStart;
					const range = sheet.getRangeByIndexes(startAbs, colIndex, runLen, 1);
					// build matrix for the run: [[v], [v], ...]
					const runMatrix = [];
					for (let k = runStart; k < runStart + runLen; k++) {
						runMatrix.push([valuesToWrite[k]]);
					}
					if (runType) {
						range.formulas = runMatrix;
					} else {
						range.values = runMatrix;
					}
				}
				// start new run
				runStart = r;
				runType = curType;
			}
		}
	}
	// leave final context.sync to caller (this function runs inside provided context)
}