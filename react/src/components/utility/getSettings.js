// Utility to read workbook settings from Office document and guarantee `columns` exists

/**
 * Return workbook settings from the Office document settings, ensuring a `columns` array exists.
 * @param {Array} defaultColumns - fallback columns to use when document has none.
 * @param {string} docKey - document settings key (default: 'workbookSettings').
 * @returns {object} settings object (may be {} with `columns` array).
 */
export function getWorkbookSettings(defaultColumns = [], docKey = 'workbookSettings') {
	try {
		if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
			const docSettings = Office.context.document.settings.get(docKey);
			if (docSettings && typeof docSettings === 'object') {
				// ensure columns exist
				if (!Array.isArray(docSettings.columns) || docSettings.columns.length === 0) {
					docSettings.columns = Array.isArray(defaultColumns) ? [...defaultColumns] : [];
				}
				// log the columns array provided
				try {
					/* eslint-disable no-console */
					console.log('getWorkbookSettings: columns ->', Array.isArray(docSettings.columns) ? [...docSettings.columns] : docSettings.columns);
					/* eslint-enable no-console */
				} catch (e) {
					// ignore logging errors
				}
				return docSettings;
			}
		}
	} catch (err) {
		/* eslint-disable no-console */
		console.warn('getWorkbookSettings: failed to read document settings', err);
		/* eslint-enable no-console */
	}
	// fallback mapping
	const fallback = { columns: Array.isArray(defaultColumns) ? [...defaultColumns] : [] };
	/* eslint-disable no-console */
	console.log('getWorkbookSettings: fallback columns ->', fallback.columns);
	/* eslint-enable no-console */
	return fallback;
}
