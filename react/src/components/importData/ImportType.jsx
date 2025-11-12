export const CanvasImport = ['student name', 'student id', 'student sis', 'course', 'course id']
export const AnthologyImport = ['student name', 'student id', 'assignment', 'score', 'comments'];
export const DropoutDetectiveImport = ['email', 'risk trend', 'course grade','course missing assignments','course zero assignments'];

// new: helper that returns an object with the detected type and the array it used to match (case-insensitive)
export function getImportType(columns = []) {
	// normalize input to lowercase trimmed strings
	const cols = (columns || []).map((c) => String(c || '').toLowerCase().trim());

	const isCanvas = CanvasImport.every((col) => cols.includes(col));
	const isAnthology = AnthologyImport.every((col) => cols.includes(col));
	const isDropoutDetective = DropoutDetectiveImport.every((col) => cols.includes(col));

	let type = 'Unknown Import';
	let matched = [];

	if (isCanvas) {
		type = 'Canvas Import';
		matched = CanvasImport;
	} else if (isAnthology) {
		type = 'Anthology Import';
		matched = AnthologyImport;
	} else if (isDropoutDetective) {
		type = 'Dropout Detective Import';
		matched = DropoutDetectiveImport;
	}

	return { type, matched };
}
