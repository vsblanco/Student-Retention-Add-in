import React, { useState } from 'react';
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

function StudentView() {
	const [activeTab, setActiveTab] = useState('assignments'); // default to 'history' tab
	const [historyData, setHistory] = useState([]); // store array returned by loadSheet
  const [assignmentData, setAssignments] = useState([]); // store array returned by loadSheet

	const loadHistory = () => {
    console.log('Loading history for', activeStudent.StudentName);
		loadSheet('Student History', 'StudentNumber', activeStudent.ID)
			.then((res) => {
				setHistory(res.data);
        console.log('Loaded history data:', res.data);
			})
			.catch((err) => {
				console.error('Failed to load Student History sheet:', err);
			});
	};
  const loadAssignments = () => {
    console.log('Loading assignments for', activeStudent.StudentName);
		loadSheet('Missing Assignments', 'gradeBookLink', activeStudent.Gradebook)
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
			return <StudentHistory history={historyData} student={activeStudent} reload={loadHistory} />;
		case 'assignments':
			return <StudentAssignments assignments={assignmentData} reload={loadAssignments} />;
		case 'details':
		default:
			return <StudentDetails student={activeStudent} />
		}
	};

	return (
		<div className="studentview-outer">
			<StudentHeader student={activeStudent} />
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