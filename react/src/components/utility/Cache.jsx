import { loadSheet } from "./ExcelAPI";

export async function loadCache(options = {}) {
	// Log start of cache load
	console.log('loadCache: start', { options });

	// Local placeholder for history data (hooks cannot be used here)
	let historyData = [];

	const activeStudentId = options.ID;
	if (activeStudentId) {
		try {
			const res = await loadSheet('Student History', 'Student identifier');
			historyData = res?.data ?? [];

			// --- NEW: persist studentHistory inside a sheetCache object in localStorage ---
			const sheetCacheKey = 'sheetCache';
			try {
				const raw = localStorage.getItem(sheetCacheKey);
				const sheetCache = raw ? JSON.parse(raw) : {};
				// store single studentHistory directly under sheetCache (replace any existing)
				sheetCache.studentHistory = historyData;
				localStorage.setItem(sheetCacheKey, JSON.stringify(sheetCache));
				console.log('loadCache: saved studentHistory to sheetCache');
			} catch (lsErr) {
				console.warn('loadCache: failed to write sheetCache to localStorage', lsErr);
			}
			// --- end new code ---
		} catch (err) {
			console.error('Failed to load Student History sheet:', err);
		}
	} else {
		console.log('loadCache: no activeStudentId provided, skipping loadSheet');
	}

	// Simulate async placeholder (can be removed when real logic is added)
	await Promise.resolve();

	// Log finish of cache load
	console.log('loadCache: finished', historyData);

	// Return the loaded data for callers to use
	return historyData;
}

// --- NEW: try to read cached studentHistory from sheetCache in localStorage ---
export async function loadSheetCache(studentId) {
	if (!studentId) return null;
  return null // temporary disable cache reading
	try {
		const raw = localStorage.getItem('sheetCache');
		if (!raw) return null;
		const sheetCache = JSON.parse(raw);
		const entry = sheetCache && sheetCache.studentHistory[studentId];
		// entry is now expected to be the history array directly
		if (entry && Array.isArray(entry)) {
			return entry;
		}
		return null;
	} catch (err) {
		console.warn('loadSheetCache: failed to read sheetCache from localStorage', err);
		return null;
	}
}