import React, { useState, useEffect, useRef, lazy, Suspense } from 'react';
import './Styling/StudentView.css';
import StudentHeader from './Parts/Header.jsx';
import StudentDetails from './Tabs/Details.jsx';
import StudentHistory, { setHistoryLoading } from './Tabs/History.jsx';
import StudentAssignments from './Tabs/Assignments.jsx';
import { onSelectionChanged, highlightRow, loadSheet, getSelectedRange } from '../utility/ExcelAPI.jsx';
import { loadCache, loadSheetCache } from '../utility/Cache.jsx';
import { isOutreachTrigger } from './Tag';

const activeStudent = {};
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
  Outreach: ['Outreach', 'Comments', 'Notes', 'Comment'],
  ProfilePicture: ['Profile Picture', 'Photo', 'Avatar']
  // You can add more aliases for other columns here
};

// add lazy SSO so its module (and any HTML it might insert) is only loaded when needed
const SSO = lazy(() => import('../utility/SSO.jsx'));

function StudentView() {
	const [activeTab, setActiveTab] = useState('history'); // default to 'history' tab
	const [historyData, setHistory] = useState([]); // store array returned by loadSheet
  	const [assignmentData, setAssignments] = useState([]); // store array returned by loadSheet
	const [activeStudentState, setActiveStudentState] = useState(activeStudent);
	const [currentUserName, setCurrentUserName] = useState(null);

	// add a ref used elsewhere in the component
	const sessionUserRef = useRef(null);

	const loadHistory = () => {
    console.log('Loading history for', activeStudentState.StudentName);

    // Trigger global/history loading skeleton before requesting the sheet
    setHistoryLoading(true);
    // loadCache is invoked once on initialization (see selection-effect)
    // Try to show cached studentHistory quickly, then always fetch fresh data
    loadSheetCache(activeStudentState.ID)
      .then((cached) => {
        if (cached) {
          setHistory(cached);
          console.log('Loaded history from sheetCache:', activeStudentState.ID);
        }
      })
      .catch((cacheErr) => {
        console.warn('loadSheetCache failed:', cacheErr);
      })
      .then(() => {
        // proceed to fetch freshest data regardless of cache result
        return loadSheet('Student History', 'StudentNumber', activeStudentState.ID);
      })
      .then((res) => {
        if (res && res.data) {
          setHistory(res.data);
          console.log('Loaded history data:', res.data);
        }
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
		console.log('Rendering tab:', activeTab);
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

		// On initial load, attempt to read the current selection and populate active student
		try {
		  const sel = await getSelectedRange(COLUMN_ALIASES);
		  if (sel && sel.success) {
		    const initialRow = sel.singleRow || (Array.isArray(sel.rows) && sel.rows[0]) || null;
		    if (initialRow) {
		      setActiveStudentState(prev => ({ ...prev, ...initialRow }));
		      console.log('Initialized active student from current selection:', initialRow);
		      // Call loadCache once during initialization with the initial row
		      try {
		        loadCache(initialRow);
		        console.log('loadCache: invoked once for initialization');
		      } catch (e) {
		        console.warn('loadCache initial invocation failed', e);
		      }
		    }
		  }
		} catch (gErr) {
		  console.warn('getSelectedRange failed to initialize selection:', gErr);
		}
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

  // Initialize currentUserName from cache/SSO on mount
  useEffect(() => {
    try {
      const cached = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
      if (cached) {
        setCurrentUserName(cached);
        sessionUserRef.current = cached;
		console.log('Loaded SSO user from cache:', cached);
        return;
      }
      if (window.SSO && typeof window.SSO.getUserName === 'function') {
        const n = window.SSO.getUserName();
        if (n) {
          setCurrentUserName(n);
          sessionUserRef.current = n;
        }
      }
    } catch (_) { /* ignore */ }
  }, []);

  // When the active student changes, reload the history for that student.
  useEffect(() => {
	// only load if there's a valid ID
	if (!activeStudentState?.ID) return;
	// Call the existing loader to update historyData
	loadHistory();
	loadAssignments();
	// eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeStudentState?.ID]);

  // Always show SSO first until currentUserName is set
  if (!currentUserName) {
    return (
      <div className="studentview-outer">
        <Suspense fallback={null}>
          <SSO onNameSelect={setCurrentUserName} />
        </Suspense>
      </div>
    );
  }

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