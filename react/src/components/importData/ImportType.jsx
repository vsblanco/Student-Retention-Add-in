import AnthologyFile from '../../assets/icons/AnthologyFile.png';
import CanvasFile from '../../assets/icons/CanvasFile.png';
import DropoutDetectiveFile from '../../assets/icons/DropoutDetectiveFile.png';

const CanvasId = 'canvas id';
const CourseId = 'course id';
export const CanvasImport = ['student sis', 'course', CourseId, 'current score']// note: not sure if to use canonical from settings or not.
// Add: mapping of detected Canvas column -> desired renamed column
export const CanvasRename = {
	'student id': CanvasId,
};

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
	let hyperLink = null; // -> new: hyperlink info when applicable
	let rename = null; // -> new: rename mapping when applicable
	let excludeFilter = null; // -> new: optional exclusion filter { column, value }

	if (isCanvas) {
		type = 'Gradebook Link';
		matched = CanvasImport;
		action = 'Update';
		icon = CanvasFile;
		rename = CanvasRename; // return the mapping to rename detected columns
		excludeFilter = { column: 'course', value: 'CAPV' }; // exclude CAPV course rows
		hyperLink = {column: 'Grade Book', // Create hyperlink to grade book
			friendlyName: 'Grade Book', 
			linkLocation: 'https://nuc.instructure.com/courses/' + CourseId + '/grades/' + CanvasId,
			parameter: [CourseId, CanvasId]
		};
	} else if (isAnthology) {
		type = 'Student List';
		matched = AnthologyImport;
		action = 'Refresh';
		icon = AnthologyFile;
	} else if (isDropoutDetective) {
		type = 'Grade';
		matched = DropoutDetectiveImport;
		action = 'Update';
		icon = DropoutDetectiveFile;
	}

	return { type, matched, action, icon, hyperLink, rename, excludeFilter };
}

