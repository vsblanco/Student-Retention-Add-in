// [2025-11-19] v5.6.1 - Added refreshSheet to getImportType return object
// Changes: 
// - Included refreshSheet in the returned object so Hybrid actions know where to refresh.

import AnthologyFile from '../../assets/icons/AnthologyLogo.png';
import CanvasFile from '../../assets/icons/CanvasLogo.png';
import DropoutDetectiveFile from '../../assets/icons/DropoutDetectiveLogo.png';
import AttendanceFile from '../../assets/icons/MyNUC-icon.png';
import MissingAssignmentFile from '../../assets/icons/MissingAssignmentsLogo.png';

const CanvasId = 'canvas id';
const CourseId = 'course id';

export const CanvasRename = {
	'student id': CanvasId,
};

// Centralized Definition for UI and Logic
export const IMPORT_DEFINITIONS = [
    {
        id: 'canvas',
        name: 'Canvas Gradebook',
        type: 'Gradebook Link',
        matchColumns: ['student sis', 'course', CourseId, 'current score'],
        action: 'Update',
        icon: CanvasFile,
        rename: CanvasRename,
        excludeFilter: { column: 'course', value: 'CAPV' },
        hyperLink: {
            column: 'Grade Book',
            friendlyName: 'Grade Book',
            linkLocation: 'https://nuc.instructure.com/courses/' + CourseId + '/grades/' + CanvasId,
            parameter: [CourseId, CanvasId]
        }
    },
    {
        id: 'anthology',
        name: 'Anthology Student List',
        type: 'Student List',
        matchColumns: ['studentname', 'studentnumber'],
        action: 'Refresh',
        icon: AnthologyFile,
    },
    {
        id: 'dropout',
        name: 'Dropout Detective',
        type: 'Grade',
        matchColumns: ['email', 'risk trend', 'course grade', 'course missing assignments', 'course zero assignments'],
        action: 'Update',
        icon: DropoutDetectiveFile,
    },
	{
        id: 'missingassignments',
        name: 'Missing Assignments Report',
        type: 'Missing Assignments',
        matchColumns: ['current grade', 'total missing','grade book'],
        action: 'Hybrid',
		refreshSheet: 'MA Test',
        icon: MissingAssignmentFile,
    },
    {
        id: 'attendance',
        name: 'MyNUC Attendance',
        type: 'LDA',
        matchColumns: ['issued id', 'date of attendance'],
        action: 'Update',
        icon: AttendanceFile,
    }
];

// Legacy exports for compatibility with other components if they use them
export const CanvasImport = IMPORT_DEFINITIONS.find(d => d.id === 'canvas').matchColumns;
export const AnthologyImport = IMPORT_DEFINITIONS.find(d => d.id === 'anthology').matchColumns;
export const DropoutDetectiveImport = IMPORT_DEFINITIONS.find(d => d.id === 'dropout').matchColumns;
export const AttendanceImport = IMPORT_DEFINITIONS.find(d => d.id === 'attendance').matchColumns;
export const MissingAssignmentsImport = IMPORT_DEFINITIONS.find(d => d.id === 'missingassignments').matchColumns;

// Function to determine import type based on provided columns

export function getImportType(columns = []) {
	const cols = (columns || []).map((c) => String(c || '').toLowerCase().trim());
    
    // Iterate through definitions to find a match
    for (const def of IMPORT_DEFINITIONS) {
        // Check if every required column exists in the input columns
        const isMatch = def.matchColumns.every(req => cols.includes(req));
        if (isMatch) {
            return {
                type: def.type,
                matched: def.matchColumns,
                action: def.action,
                icon: def.icon,
                hyperLink: def.hyperLink || null,
                rename: def.rename || null,
                excludeFilter: def.excludeFilter || null,
                refreshSheet: def.refreshSheet || null // Added this line
            };
        }
    }

    // Default fallback
	return { 
        type: 'Standard', 
        matched: [], 
        action: 'Refresh', 
        icon: null, 
        hyperLink: null, 
        rename: null, 
        excludeFilter: null,
        refreshSheet: null
    };
}