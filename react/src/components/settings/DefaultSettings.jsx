import React from 'react';
import importIcon from '../../assets/icons/import-icon.png';
import createLdaIcon from '../../assets/icons/create-lda-icon.png';
import emailIcon from '../../assets/icons/email-icon.png';

export const defaultColumns =[
	{ name: 'Assigned', alias: ['advisor'], static: true },
	{ name: 'Student Name', alias: ['Student']},
	{ name: 'Student ID', alias: ['Student Number','Student identifier'], hidden: true },
	{ name: 'Gender', alias: ['Sex'], hidden: true},
	{ name: 'Phone', alias: ['Phone Number','Contact Number']},
	{ name: 'Other Phone', alias: ['Second Phone', 'Alt Phone']},
	{ name: 'StudentEmail', alias: ['Email', 'Student Email'], hidden: true},
	{ name: 'PersonalEmail', alias: ['Other Email'], hidden: true},
	{ name: 'Gradebook', alias: ['Gradelink','gradeBookLink'], static: true, format: ['=HYPERLINK']},
	{ name: 'Grade', alias: ['Current Grade', 'Grade Level', 'Score']  },
	{ name: 'LDA', alias: ['Last Date of Attendance'], format: ['MM/DD/YYYY']},
	{ name: 'Days Out', alias: ['Days Out'], format: ['G-Y-R Color Scale']},
	{ name: 'Photo', alias: ['pfp', 'profile photo'], static: true, hidden: true },
];

export const formatOptions = [
	{ option: 'Conditional Formatting', type: ['G-Y-R Color Scale', 'R-Y-G Color Scale'] },
	{ option: 'Function', type: ['=HYPERLINK', '=XLOOKUP'] },
	{ option: 'Date Format', type: ['MM/DD/YYYY', 'DD/MM/YYYY', 'YYYY-MM-DD'] },
];

export const defaultWorkbookSettings = [
	
	{
		id: 'masterListName',
		label: 'Master List Name',
		type: 'text',
		defaultValue: 'Master List',
		section: 'Import Data',
		description: 'The name of the master list to import data from.'
	},
	{
		id: 'columnsToImport',
		label: 'Columns to Import',
		type: 'array',
		defaultValue: defaultColumns,
		section: 'Import Data',
		description: 'List of columns to import from the master list. Each item is an object: { name, alias, static } where alias may be a string or an array of strings and static is a boolean to prevent edits.'
	},
	{
		id: 'Format',
		label: 'Format',
		type: 'editableArray',
		choices: defaultColumns,
		options: formatOptions,
		defaultValue: [],
		section: 'Import Data',
		description: 'List of columns to format in the master list.'
	},
	{
		id: 'daysOut',
		label: 'Days Out',
		type: 'number',
		defaultValue: 5,
        section: 'Create LDA',
		description: 'Number of days out threshold.'
	},
	{
		id: 'includeFailingList',
		label: 'Include Failing List',
		type: 'boolean',
		defaultValue: false, // No
        section: 'Create LDA',
		description: 'Whether to include the failing student list.'
	},
	{
		id: 'includeLdatTag',
		label: 'Include LDA Tag',
		type: 'boolean',
		defaultValue: true, // Yes
        section: 'Create LDA',
		description: 'Whether to include the LDA tag.'
	},
		{
		id: 'includeDncTag',
		label: 'Include DNC Tag',
		type: 'boolean',
		defaultValue: true, // Yes
        section: 'Create LDA',
		description: 'Whether to include the Do Not Contact students.'
	},
		{
		id: 'hiddenColumns',
		label: 'Hidden Columns',
		type: 'selections',
		choices: defaultColumns,
		defaultValue: [],
		key: 'hidden',
		section: 'Create LDA',
		description: 'Defines the structure and visibility of the columns in the master list.'
	},
	{
		id: 'powerAutomateFlowUrl',
		label: 'Power Automate Flow URL',
		type: 'text',
		defaultValue: '',
		placeholderValue: 'URL to send HTTP Request to trigger email flow',
		section: 'Personalized Emails',
		description: 'The URL of the Power Automate flow to trigger.'
	}
	
];

// create a small SVG avatar and encode as data URI for the default profile picture
const _defaultAvatarSvg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'>
	<rect fill='%23e5e7eb' width='100' height='100' rx='16'/>
	<circle fill='%239ca3af' cx='50' cy='36' r='18'/>
	<path fill='%239ca3af' d='M20 84c0-14 12-26 30-26s30 12 30 26' />
	</svg>`;
const defaultProfilePicture = `data:image/svg+xml;utf8,${encodeURIComponent(_defaultAvatarSvg)}`;

export const defaultUserSettings = [
	{
		id: 'theme',
		label: 'Theme',
		type: 'select',
		defaultValue: 'light',
		options: ['light', 'dark', 'system'],
		description: 'User interface theme preference.'
	},
	{
		id: 'profilePicture',
		label: 'Profile Picture',
		type: 'image',
		defaultValue: defaultProfilePicture,
		description: 'Default profile avatar (data URI).'
	}
];

// export lucide-react icons keyed by section name
export const sectionIcons = {
	'Import Data': <img src={importIcon} alt="Import" style={{ width: 16, height: 16 }} aria-hidden="true" />,
	'Create LDA': <img src={createLdaIcon} alt="Create LDA" style={{ width: 16, height: 16 }} aria-hidden="true" />,
	'Personalized Emails': <img src={emailIcon} alt="Emails" style={{ width: 16, height: 16 }} aria-hidden="true" />
};

// Optionally default export for convenience
export default {
	defaultWorkbookSettings,
	defaultUserSettings
};
