import AnthologyFile from '../../assets/icons/AnthologyFile.png';
import CanvasFile from '../../assets/icons/CanvasFile.png';
import DropoutDetectiveFile from '../../assets/icons/DropoutDetectiveFile.png';

export const CanvasImport = ['student sis', 'course', 'course id', 'grade']// note: not sure if to use canonical from settings or not.
export const AnthologyImport = ['studentname', 'studentnumber']; // for ssome reason it has to be lowercase and no spaces
export const DropoutDetectiveImport = ['email', 'risk trend', 'course grade','course missing assignments','course zero assignments'];

// new: helper that returns an object with the detected type, the array it used to match (case-insensitive),
// and an "action" indicating whether to use Update or Refresh
export function getImportType(columns = []) {
	// normalize input to lowercase trimmed strings
	const cols = (columns || []).map((c) => String(c || '').toLowerCase().trim());

	const isCanvas = CanvasImport.every((col) => cols.includes(col));
	const isAnthology = AnthologyImport.every((col) => cols.includes(col));
	const isDropoutDetective = DropoutDetectiveImport.every((col) => cols.includes(col));

	let type = 'Standard';
	let matched = [];
	let action = 'Refresh';
	let icon = null; // -> new: icon to return
	let urlBuilder = null; // -> now an object { build: Function, to: string } when applicable

	if (isCanvas) {
		type = 'Gradebook Link';
		matched = CanvasImport;
		action = 'Update';
		icon = CanvasFile;
		// return both the builder and the target column name ("to") so callers know where to place the URL
		urlBuilder = { build: createCanvasGradebookUrl, from: requiredGradebookParams, to: GradebookColumn };
	} else if (isAnthology) {
		type = 'Student Population';
		matched = AnthologyImport;
		action = 'Refresh';
		icon = AnthologyFile;
	} else if (isDropoutDetective) {
		type = 'Grade';
		matched = DropoutDetectiveImport;
		action = 'Update';
		icon = DropoutDetectiveFile;
	}

	return { type, matched, action, icon, urlBuilder };
}

// New: small helper to build a Canvas gradebook URL.
// - courseId and studentId are required (will be stringified & encoded).
// - baseUrl defaults to https://nuc.instructure.com and trailing slash is handled.
// Example: createCanvasGradebookUrl('123', '456') -> 'https://nuc.instructure.com/courses/123/grades/456'
const studentIdParam = 'student id';
const courseIdParam = 'course id';
const courseParam = 'course';
//----------------------------------
const requiredGradebookParams = [studentIdParam, courseIdParam, courseParam];
const GradebookColumn = 'Grade Book';
const CanvasBaseUrl = 'https://nuc.instructure.com';
const IgnoreCapv = 'CAPV';
//----------------------------------
export function createCanvasGradebookUrl(from = [], baseUrl = CanvasBaseUrl) {
	// Helper: build URL for a single row (row can be an Array or an Object)
	function buildForRow(row) {
		// normalize row to an array [studentId, courseId, courseName]
		let arr = [];
		if (Array.isArray(row)) {
			arr = row.slice(0, 3);
		} else if (row && typeof row === 'object') {
			// try to pick values from object using variations of required params
			arr = requiredGradebookParams.map((p) => {
				const raw = String(p || '');
				const noSpace = raw.replace(/\s+/g, '').toLowerCase();
				// prefer explicitly named keys without spaces (e.g., studentid), then original, then lowercase
				return row[noSpace] ?? row[raw] ?? row[raw.toLowerCase()] ?? '';
			});
		} else {
			arr = [];
		}

		const studentId = arr[0] != null ? String(arr[0]) : studentIdParam;
		let courseId = arr[1] != null ? String(arr[1]) : courseIdParam;
		const course = arr[2] != null ? String(arr[2]) : courseParam;

		// If the course string (or courseId) contains "CAPV" (case-insensitive), skip the whole process.
		if (String(course || '').toUpperCase().includes(IgnoreCapv) || String(courseId || '').toUpperCase().includes(IgnoreCapv)) {
			return null; // caller can treat null as "skip"
		}
		// normalize and protect against accidental trailing slash on baseUrl
		const base = String(baseUrl || '').replace(/\/+$/, '');
		const c = encodeURIComponent(String(courseId));
		const s = encodeURIComponent(String(studentId));
		return `${base}/courses/${c}/grades/${s}`;
	}

	// If caller provided an array-of-rows (arrays or objects), handle both cases.
	if (Array.isArray(from) && from.length > 0 && (Array.isArray(from[0]) || typeof from[0] === 'object')) {
		const isObjectArray = typeof from[0] === 'object' && !Array.isArray(from[0]);
		if (isObjectArray) {
			// Return an array of objects with the GradebookColumn property added
			return from.map((obj) => {
				const url = buildForRow(obj);
				// shallow copy and add the Grade Book column (url or null)
				return { ...obj, [GradebookColumn]: url };
			});
		}
		// existing behavior for array-of-arrays: return array of urls (or nulls)
		return from.map((r) => buildForRow(r));
	}

	// If a single object row is provided, return the object with the Grade Book property added
	if (from && typeof from === 'object' && !Array.isArray(from)) {
		const url = buildForRow(from);
		return { ...from, [GradebookColumn]: url };
	}

	// Otherwise treat `from` as a single row (array or primitive) and return a single url or null
	return buildForRow(from);
}
