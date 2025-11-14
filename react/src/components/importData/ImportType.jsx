import AnthologyFile from '../../assets/icons/AnthologyFile.png';
import CanvasFile from '../../assets/icons/CanvasFile.png';
import DropoutDetectiveFile from '../../assets/icons/DropoutDetectiveFile.png';

export const CanvasImport = ['student name', 'student sis', 'course', 'course id']
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

	if (isCanvas) {
		type = 'Gradebook Link';
		matched = CanvasImport;
		action = 'Update';
		icon = CanvasFile;
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

	return { type, matched, action, icon };
}
