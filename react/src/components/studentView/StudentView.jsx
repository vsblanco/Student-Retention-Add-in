import React, { useState, useEffect } from 'react';
import './Styling/StudentView.css';
import StudentHeader from './Parts/Header.jsx';
import StudentDetails from './Tabs/Details.jsx';
import StudentHistory from './Tabs/History.jsx';
import StudentAssignments from './Tabs/Assignments.jsx';
import { onSelectionChanged, highlightRow, loadSheet } from '../utility/ExcelAPI.jsx';
import { isOutreachTrigger } from './Tag';

const activeStudent = {
  StudentName: 'Saintil, Dominique',
  ID: 1612676588,
  Grade: '10',
  Gradebook: 'https://nuc.instructure.com/courses/103987/grades/168591',
};
export const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'Student Number','Student identifier'],
  Gender: ['Gender'],
  Phone: ['Phone Number', 'Contact'],
  CreatedBy: ['Created By', 'Author', 'Advisor'],
  OtherPhone: ['Other Phone', 'Alt Phone'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor'],
  ExpectedStartDate: ['Expected Start Date', 'Start Date','ExpStartDate'],
  Grade: ['Current Grade', 'Grade %', 'Grade'],
  LDA: ['Last Date of Attendance', 'LDA'],
  DaysOut: ['Days Out'],
  Gradebook: ['Gradebook','gradeBookLink'],
  MissingAssignments: ['Missing Assignments', 'Missing'],
  Outreach: ['Outreach', 'Comments', 'Notes', 'Comment']
  // You can add more aliases for other columns here
};

function StudentView() {
	const [activeTab, setActiveTab] = useState('assignments'); // default to 'history' tab
	const [historyData, setHistory] = useState([]); // store array returned by loadSheet
  	const [assignmentData, setAssignments] = useState([]); // store array returned by loadSheet
	const [activeStudentState, setActiveStudentState] = useState(activeStudent);

	const loadHistory = () => {
    console.log('Loading history for', activeStudentState.StudentName);
		loadSheet('Student History', 'StudentNumber', activeStudentState.ID)
			.then((res) => {
				setHistory(res.data);
        console.log('Loaded history data:', res.data);
			})
			.catch((err) => {
				console.error('Failed to load Student History sheet:', err);
			});
	};
  const loadAssignments = () => {
    console.log('Loading assignments for', activeStudentState.StudentName);
		loadSheet('Missing Assignments', 'gradeBookLink', activeStudentState.Gradebook)
			.then((res) => {
				setAssignments(res.data);
        console.log('Loaded assignments data:', res.data);
			})
			.catch((err) => {
				console.error('Failed to load Student Assignments sheet:', err);
			});
	};

	const renderActiveTab = () => {
		switch (activeTab) {
		case 'history':
			return <StudentHistory history={historyData} student={activeStudentState} reload={loadHistory} />;
		case 'assignments':
			return <StudentAssignments assignments={assignmentData} reload={loadAssignments} />;
		case 'details':
		default:
			return <StudentDetails student={activeStudentState} />
		}
	};

  // Register Excel selection-changed handler to log selection details.
  useEffect(() => {
	let handlerRef = null;
	(async () => {
	  try {
		// pass COLUMN_ALIASES so headers are canonicalized in the callback
		handlerRef = await onSelectionChanged(({ address, values, data }) => {
		  console.log('Excel selection changed:', {data});
		  // merge selected row data into active student state
		  setActiveStudentState(prev => ({ ...prev, ...data }));
		}, COLUMN_ALIASES);
	  } catch (err) {
		console.error('Failed to register Excel selection handler:', err);
	  }
	})();

	// Cleanup: unregister the handler when the component unmounts
	return () => {
	  if (handlerRef && typeof handlerRef.remove === 'function') {
		try {
		  handlerRef.remove();
		} catch (err) {
		  console.warn('Failed to remove selection handler:', err);
		}
	  }
	};
  }, []);

	return (
		<div className="studentview-outer">
			<StudentHeader student={activeStudentState} />
			<div className="studentview-tabs">
				<button
					type="button"
					className={`studentview-tab ${activeTab === 'details' ? 'active' : ''}`}
					onClick={() => setActiveTab('details')}
				>
					Details
				</button>
				<button
					type="button"
					className={`studentview-tab ${activeTab === 'history' ? 'active' : ''}`}
					onClick={() => {
            loadHistory();
						setActiveTab('history');
					}}
				>
					History
				</button>
				<button
					type="button"
					className={`studentview-tab ${activeTab === 'assignments' ? 'active' : ''}`}
					onClick={() => {
            loadAssignments();
						setActiveTab('assignments');
					}}
				>
					Assignments
				</button>
			</div>

			<div className="studentview-tab-content">
				{renderActiveTab()}
			</div>
		</div>
	);
}

StudentView.displayName = 'StudentView';

export default React.memo(StudentView);