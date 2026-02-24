import React, { useState, useEffect } from 'react';
import { Info } from 'lucide-react'; // removed X + Plus imports (moved to SettingsModal)
import '../studentView/Styling/StudentView.css'; // add StudentView tab styles
import { defaultUserSettings, defaultWorkbookSettings, defaultColumns, sectionIcons } from './DefaultSettings'; // added: import defaults + defaultColumns
import SettingsModal from './SettingsModal'; // new: modal component
import WorkbookSettingsModal from './WorkbookSettingsModal'; // <-- ADDED
import PowerAutomateConfigModal from './PowerAutomateConfigModal'; // <-- ADDED: Power Automate config modal
import LicenseChecker from '../utility/LicenseChecker'; // <-- License checker (requires Graph API)
import UserInfoDisplay from '../utility/UserInfoDisplay'; // <-- User info from token (no API needed)
import About from '../about/About'; // <-- ADDED: Import About component for Help tab

const Settings = ({ user, accessToken, onReady }) => { // <-- ADDED accessToken and onReady props
	const [activeTab, setActiveTab] = useState('workbook');

	// initialize user settings state from defaults
	const [userSettingsState, setUserSettingsState] = useState(() =>
		defaultUserSettings.reduce((acc, s) => {
			acc[s.id] = s.defaultValue;
			return acc;
		}, {})
	);

	// initialize workbook settings state from defaults
	const [workbookSettingsState, setWorkbookSettingsState] = useState(() =>
		defaultWorkbookSettings.reduce((acc, s) => {
			acc[s.id] = s.defaultValue;
			return acc;
		}, {})
	);

	// modal state to edit array-type settings
	const [modalOpen, setModalOpen] = useState(false);
	const [modalSetting, setModalSetting] = useState(null);
	const [modalArray, setModalArray] = useState([]);
	const [modalUpdater, setModalUpdater] = useState(null);

	// new: workbook inspector modal state
	const [workbookModalOpen, setWorkbookModalOpen] = useState(false);

	// Power Automate config modal state
	const [powerAutomateModalOpen, setPowerAutomateModalOpen] = useState(false);

	// Download CSV state
	const [downloadingCsv, setDownloadingCsv] = useState(false);

	// master list headers read from the active workbook (populated when columns modal opens)
	const [masterListHeaders, setMasterListHeaders] = useState(null);

	// set of Master List headers that are NEW (not seen when the user last saved columns)
	const [newMasterListHeaders, setNewMasterListHeaders] = useState(null);

	// set of Master List headers whose data columns are entirely empty/blank
	const [blankColumns, setBlankColumns] = useState(null);

	// read headers from the "Master List" worksheet via Excel API
	// Returns { headers, blankColumns } where blankColumns is a Set of header names whose data columns are entirely empty
	const readMasterListHeaders = async () => {
		try {
			if (typeof window !== 'undefined' && window.Excel && Excel.run) {
				const result = await Excel.run(async (context) => {
					const sheet = context.workbook.worksheets.getItemOrNullObject('Master List');
					await context.sync();
					if (sheet.isNullObject) return null;
					const used = sheet.getUsedRangeOrNullObject();
					used.load('values');
					await context.sync();
					if (used.isNullObject || !used.values || used.values.length === 0) return null;
					const allValues = used.values;
					const headers = allValues[0].map(v => (v == null ? '' : String(v).trim())).filter(Boolean);
					// detect blank columns: columns where every data row (rows 1+) is empty/null
					const blankSet = new Set();
					headers.forEach((header, colIdx) => {
						let allEmpty = true;
						for (let row = 1; row < allValues.length; row++) {
							const cell = allValues[row][colIdx];
							if (cell != null && String(cell).trim() !== '') {
								allEmpty = false;
								break;
							}
						}
						// only mark as blank if there are data rows; a header-only sheet is not "blank"
						if (allEmpty && allValues.length > 1) {
							blankSet.add(header);
						}
					});
					return { headers, blankColumns: blankSet };
				});
				return result;
			}
		} catch (err) {
			console.warn('Failed to read Master List headers', err);
		}
		return null;
	};

	// Download "Student History" sheet as CSV
	const downloadHistoryCsv = async () => {
		if (downloadingCsv) return;
		setDownloadingCsv(true);
		try {
			if (typeof window === 'undefined' || !window.Excel || !Excel.run) {
				throw new Error('Excel API not available');
			}
			const csvContent = await Excel.run(async (context) => {
				const sheet = context.workbook.worksheets.getItemOrNullObject('Student History');
				await context.sync();
				if (sheet.isNullObject) throw new Error('Student History sheet not found');
				const used = sheet.getUsedRangeOrNullObject();
				used.load('values');
				await context.sync();
				if (used.isNullObject || !used.values || used.values.length === 0) {
					throw new Error('Student History sheet is empty');
				}
				// Build CSV from the 2D values array
				return used.values.map(row =>
					row.map(cell => {
						const val = cell == null ? '' : String(cell);
						// Escape cells containing commas, quotes, or newlines
						if (val.includes(',') || val.includes('"') || val.includes('\n')) {
							return `"${val.replace(/"/g, '""')}"`;
						}
						return val;
					}).join(',')
				).join('\n');
			});
			// Trigger download
			const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
			const url = URL.createObjectURL(blob);
			const a = document.createElement('a');
			a.href = url;
			a.download = 'Student History.csv';
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			URL.revokeObjectURL(url);
		} catch (err) {
			console.error('Failed to download Student History CSV:', err);
			alert(err.message || 'Failed to download Student History');
		} finally {
			setDownloadingCsv(false);
		}
	};

	const openArrayModal = async (setting, currentValue, updater) => {
		// if this is the columns setting, read master list headers first
		if (setting.id === 'columns') {
			const result = await readMasterListHeaders();
			const headers = result ? result.headers : null;
			setMasterListHeaders(headers); // null means no master list found (show all)
			setBlankColumns(result ? result.blankColumns : null);

			// detect new headers by comparing current ML headers against lastKnownHeaders
			const lastKnown = workbookSettingsState.lastKnownHeaders;
			if (headers && Array.isArray(lastKnown) && lastKnown.length > 0) {
				const stripStr = s => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
				const knownSet = new Set(lastKnown.map(h => stripStr(h)));
				const newSet = new Set();
				headers.forEach(h => {
					if (!knownSet.has(stripStr(h))) newSet.add(h);
				});
				setNewMasterListHeaders(newSet.size > 0 ? newSet : null);
			} else {
				// first time opening or no lastKnownHeaders — nothing is "new"
				setNewMasterListHeaders(null);
			}
		} else {
			setMasterListHeaders(null);
			setNewMasterListHeaders(null);
			setBlankColumns(null);
		}

		setModalSetting(setting);
		// normalize items to objects: { name: string, alias: string[], edit: 'name'|'alias', static: boolean }
		const normalized = Array.isArray(currentValue)
			? currentValue.map(it => {
					if (typeof it === 'string') return { name: it, alias: [], edit: 'name', static: false };
					if (it && typeof it === 'object') {
						return {
							name: it.name ?? '',
							alias: Array.isArray(it.alias)
								? it.alias
								: it.alias
									? String(it.alias).split(',').map(s => s.trim()).filter(Boolean)
									: [],
							edit: 'name',
							static: !!it.static
						};
					}
					return { name: '', alias: [], edit: 'name', static: false };
			  })
			: [];
		setModalArray(normalized);
		// store updater as a function reference
		setModalUpdater(() => updater);
		setModalOpen(true);
	};

	const closeModal = () => {
		setModalOpen(false);
		setModalSetting(null);
		setModalArray([]);
		setModalUpdater(null);
	};

	// modal save: accept optional newArray provided by the modal to avoid reading stale state
	const saveModal = (providedArray) => {
		if (modalSetting && typeof modalUpdater === 'function') {
			// prefer the array passed from the modal (freshly computed) otherwise use parent's modalArray state
			const sourceArray = Array.isArray(providedArray) ? providedArray : modalArray;
			// ensure we send a clean array of objects: { name, alias: string[], static: boolean, ...options }
			const cleaned = (Array.isArray(sourceArray) ? sourceArray : []).map(it => ({
				name: String(it?.name ?? it?.column ?? '').trim(),
				alias: Array.isArray(it?.alias) ? it.alias.map(a => String(a).trim()).filter(Boolean) : [],
				static: !!it?.static,
				// preserve any other option fields the editor wrote under `options`
				...((it && it.options && typeof it.options === 'object') ? it.options : {})
			}));
			modalUpdater(modalSetting.id, cleaned);

			// when saving columns, snapshot current Master List headers so we can detect new ones later
			if (modalSetting.id === 'columns' && Array.isArray(masterListHeaders)) {
				updateWorkbookSetting('lastKnownHeaders', masterListHeaders);
				setNewMasterListHeaders(null); // clear "new" badges after save
			}
		}
		// close modal via parent helper
		closeModal();
	};

	// Persist/load helpers for Office document (Workbook) settings
	const DOC_KEY = 'workbookSettings';

	const loadWorkbookSettingsFromDocument = () => {
		try {
			// ensure Office.js is available and document settings exist
			if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				const docSettings = Office.context.document.settings.get(DOC_KEY);
				return docSettings || null;
			}
		} catch (err) {
			console.warn('Failed to read document settings', err);
		}
		return null;
	};

	const saveWorkbookSettingsToDocument = (mapping) => {
		try {
			if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				Office.context.document.settings.set(DOC_KEY, mapping);
				Office.context.document.settings.saveAsync(result => {
					if (result && result.status !== Office.AsyncResultStatus.Succeeded) {
						console.warn('Failed to save workbook settings to document', result.error);
					}
				});
			}
		} catch (err) {
			console.warn('Failed to save document settings', err);
		}
	};

	// one-time effect: ensure document has a workbook settings key; load it if present
	useEffect(() => {
		const existing = loadWorkbookSettingsFromDocument();
		if (existing && typeof existing === 'object') {
			// if columns are missing or empty, inject defaultColumns (one-time import)
			const needsColumns = !Array.isArray(existing.columns) || existing.columns.length === 0;
			const merged = { ...existing };
			if (needsColumns) {
				merged.columns = [...defaultColumns];
				// persist the merged mapping back to document settings
				saveWorkbookSettingsToDocument(merged);
			}
			setWorkbookSettingsState(prev => ({ ...prev, ...merged }));
			return;
		}

		// no settings in document: create mapping from defaultWorkbookSettings and persist once
		const initial = defaultWorkbookSettings.reduce((acc, s) => {
			acc[s.id] = s.defaultValue;
			return acc;
		}, {});
		// ensure columns are populated from defaultColumns on initial import
		if (!Array.isArray(initial.columns) || initial.columns.length === 0) {
			initial.columns = [...defaultColumns];
		}
		setWorkbookSettingsState(initial);
		saveWorkbookSettingsToDocument(initial);
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []); // run once on mount

	// Signal that Settings is ready
	useEffect(() => {
		if (onReady) {
			onReady();
		}
	}, [onReady]);

	// helper to update a workbook setting (also persist to document)
	const updateWorkbookSetting = (id, value) => {
		setWorkbookSettingsState(prev => {
			const next = { ...prev, [id]: value };
			// persist mapping to document settings
			saveWorkbookSettingsToDocument(next);
			return next;
		});
	};

	// helper to update a setting
	const updateSetting = (id, value) => {
		setUserSettingsState(prev => ({ ...prev, [id]: value }));
	};

	// render settings grouped by optional `section` property
	const renderSettingsControls = (settings, state, updater, idPrefix = '') => {
		// group by section
		const sections = {};
		const unsectioned = [];
		settings.forEach(s => {
			if (s.section) {
				(sections[s.section] ||= []).push(s);
			} else {
				unsectioned.push(s);
			}
		});

		const renderRow = setting => {
			const cur = state[setting.id];
			const inputId = `${idPrefix}${setting.id}`;
			return (
				<div key={setting.id} style={{ display: 'grid', gridTemplateColumns: '1fr auto', alignItems: 'center', gap: 8, width: '100%', boxSizing: 'border-box' }}>
					{/* label area: stays in the left column and truncates if too long */}
					<div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0, overflow: 'hidden' }}>
						<div style={{ fontSize: 14, fontWeight: 600, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
							{setting.label}
						</div>
						{/* info icon kept compact and closer to the label */}
						{setting.description && (
							<button
								type="button"
								title={setting.description}
								aria-label={`${setting.label} description`}
								style={{
									marginLeft: 1,
									padding: 1,
									border: 'none',
									background: 'transparent',
									color: '#6b7280',
									cursor: 'pointer',
									display: 'inline-flex',
									alignItems: 'center',
								}}
							>
								<Info size={14} />
							</button>
						)}
					</div>

					{/* input/control container: reduced gap so control sits nearer the label */}
					<div style={{ display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'flex-start' }}>
						{(() => {
							if (setting.type === 'boolean') {
								// Tailwind-like toggle: accessible hidden checkbox + visual track & knob
								return (
									<label style={{ display: 'inline-flex', alignItems: 'center', gap: 1, cursor: 'pointer' }}>
										<span style={{ position: 'relative', width: 44, height: 24, display: 'inline-block' }}>
											<input
												type="checkbox"
												checked={!!cur}
												onChange={e => updater(setting.id, e.target.checked)}
												aria-label={setting.label}
												style={{ position: 'absolute', opacity: 0, width: 0, height: 0 }}
											/>
											{/* track */}
											<span style={{
												position: 'absolute',
												inset: 0,
												borderRadius: 9999,
												background: !!cur ? '#4f46e5' : '#e5e7eb',
												transition: 'background-color 160ms linear',
												padding: 2,
												boxSizing: 'border-box'
											}} />
											{/* knob */}
											<span style={{
												position: 'absolute',
												top: 2,
												left: !!cur ? 22 : 2,
												width: 20,
												height: 20,
												borderRadius: '50%',
												background: '#fff',
												boxShadow: '0 1px 2px rgba(0,0,0,0.2)',
												transition: 'left 160ms linear',
											}} />
										</span>
										<span style={{ fontSize: 13 }}>{cur ? 'On' : 'Off'}</span>
									</label>
								);
							}

							if (setting.type === 'select' && Array.isArray(setting.options)) {
								return (
									<select value={cur ?? ''} onChange={e => updater(setting.id, e.target.value)} style={{ padding: '6px 10px', borderRadius: 6, minWidth: 120 }} aria-label={setting.label}>
										{setting.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
									</select>
								);
							}

							if (setting.type === 'array') {
								// show disabled Configure button with lock indicator when locked
								if (setting.locked) {
									return (
										<button
											onClick={() => {}}
											style={{
												padding: '6px 10px',
												borderRadius: 6,
												background: '#f9fafb',
												border: '1px solid #e6e7eb',
												cursor: 'not-allowed',
												display: 'inline-flex',
												alignItems: 'center',
												gap: 8,
												color: '#6b7280'
											}}
											aria-label={`Configure`}
											aria-disabled="true"
											title="This setting is locked"
										>
											{/* small lock icon */}
											<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', color: '#9ca3af' }}>
												<rect x="3" y="11" width="18" height="10" rx="2" stroke="currentColor" strokeWidth="1.2" />
												<path d="M7 11V8a5 5 0 0 1 10 0v3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
											</svg>
											<span style={{ fontSize: 13 }}>Locked</span>
										</button>
									);
								}
								return (
									<button
										onClick={() => openArrayModal(setting, cur, updater)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'powerautomate') {
								return (
									<button
										onClick={() => setPowerAutomateModalOpen(true)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure Power Automate`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'action') {
								const isDownloading = setting.id === 'downloadHistoryCsv' && downloadingCsv;
								return (
									<button
										onClick={() => {
											if (setting.id === 'downloadHistoryCsv') downloadHistoryCsv();
										}}
										disabled={isDownloading}
										style={{
											padding: '6px 10px',
											borderRadius: 6,
											background: isDownloading ? '#e5e7eb' : '#f3f4f6',
											border: '1px solid #e6e7eb',
											cursor: isDownloading ? 'not-allowed' : 'pointer',
											display: 'inline-flex',
											alignItems: 'center',
											gap: 6,
											fontSize: 13,
										}}
										aria-label={setting.label}
									>
										{/* download icon */}
										<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
											<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
											<polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
											<line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
										</svg>
										{isDownloading ? 'Downloading...' : 'Download'}
									</button>
								);
							}

							if (setting.type === 'image') {
								const fileId = `file-${inputId}`;
								return (
									<div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
										<input id={fileId} type="file" accept="image/*" style={{ display: 'none' }} onChange={e => {
											const f = e.target.files && e.target.files[0];
											if (f) {
												const r = new FileReader();
												r.onload = () => updater(setting.id, r.result);
												r.readAsDataURL(f);
											}
										}} aria-label={`Upload ${setting.label}`} />
										<label htmlFor={fileId} title="Upload" style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 32, height: 32, borderRadius: 6, background: '#f3f4f6', cursor: 'pointer', border: '1px solid #e5e7eb', color: '#374151' }}>
											<svg width="18" height="18" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
												<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
												<polyline points="17 8 12 3 7 8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
												<line x1="12" y1="3" x2="12" y2="15" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
											</svg>
										</label>
										<button onClick={() => updater(setting.id, setting.defaultValue)} style={{ padding: '6px 8px', borderRadius: 6 }} title="Reset to default">Reset</button>
									</div>
								);
							}

							if (setting.type === 'number') {
								return (
									<input
										type="number"
										value={cur === undefined || cur === null ? '' : cur}
										onChange={e => updater(setting.id, e.target.value === '' ? '' : Number(e.target.value))}
										style={{
											padding: '6px 8px',
											borderRadius: 6,
											width: 72,
											boxSizing: 'border-box',
											border: '1px solid #e6e7eb', /* very subtle outline */
											background: '#fff',
											boxShadow: 'inset 0 0 0 1px rgba(0,0,0,0.02)',
											transition: 'box-shadow 120ms ease, border-color 120ms ease'
										}}
										placeholder={setting.placeholderValue ?? ''}
										aria-label={setting.label}
									/>
								);
							}

							return (
								<input
									type="text"
									value={cur ?? ''}
									onChange={e => updater(setting.id, e.target.value)}
									style={{
										padding: '6px 10px',
										borderRadius: 6,
										border: '1px solid #e6e7eb', /* very subtle outline */
										background: '#fff',
										boxShadow: 'inset 0 0 0 1px rgba(0,0,0,0.02)',
										transition: 'box-shadow 120ms ease, border-color 120ms ease',
										minWidth: 0
									}}
									placeholder={setting.placeholderValue ?? ''}
									aria-label={setting.label}
								/>
							);
						})()}
					</div>
				</div>
 			);
 		};

		return (
			<div style={{ display: 'grid', gap: 8 }}>
				{Object.keys(sections).map(name => (
					<div key={name}>
						<h3 style={{ margin: '2px 0', fontSize: 15, fontWeight: 700, backgroundColor: '#eaeaea', padding: '2px 6px', borderRadius: 6 }}>
							<span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
								{sectionIcons[name] && (
									<span aria-hidden="true" style={{ display: 'inline-flex', width: 16, height: 16, alignItems: 'center' }}>
										{sectionIcons[name]}
									</span>
								)}
								<span>{name}</span>
							</span>
						</h3>
						<div style={{ padding: 8, border: '1px solid #f3f4f6', borderRadius: 6, background: '#fafafa', display: 'grid', gap: 8 }}>
							{sections[name].map(renderRow)}
						</div>
					</div>
				))}
				{unsectioned.map(renderRow)}
			</div>
		);
	};

	// handle button actions per setting type
	const handleSettingButton = (setting) => {
		const cur = userSettingsState[setting.id];
		if (setting.type === 'boolean') {
			updateSetting(setting.id, !cur);
		} else if (setting.type === 'select' && Array.isArray(setting.options)) {
			const idx = setting.options.indexOf(cur);
			const next = setting.options[(idx + 1) % setting.options.length];
			updateSetting(setting.id, next);
		} else if (setting.type === 'image') {
			// reset to default image on click
			updateSetting(setting.id, setting.defaultValue);
		} else if (setting.type === 'number') {
			// simple increment for demo
			updateSetting(setting.id, (typeof cur === 'number' ? cur + 1 : setting.defaultValue));
		} else {
			// fallback: toggle between default and empty
			updateSetting(setting.id, cur === setting.defaultValue ? '' : setting.defaultValue);
		}
	};

	return (
		<div
			className="settings-placeholder"
			role="region"
			aria-label="Settings"
			style={{ position: 'relative' }} // added: anchor for absolute avatar
		>
			<div
				className="settings-header"
				style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}
			>
				<div style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
					<h2 className="text-2xl text-slate-800 font-bold tracking-tight" style={{ margin: 0, paddingLeft: 12 }}>Settings</h2>
					<p className="text-slate-400 text-sm mt-1">Manage your workbook and user preferences</p>

					{/* Tabs - use same classes as StudentView for identical styling */}
					<div className="studentview-tabs" role="tablist" aria-label="Settings tabs" style={{ marginBottom: 8 }}>
						<button
							role="tab"
							aria-selected={activeTab === 'workbook'}
							onClick={() => setActiveTab('workbook')}
							className={`studentview-tab ${activeTab === 'workbook' ? 'active' : ''}`}
						>
							Workbook
						</button>
						<button
							role="tab"
							aria-selected={activeTab === 'user'}
							onClick={() => setActiveTab('user')}
							className={`studentview-tab ${activeTab === 'user' ? 'active' : ''}`}
						>
							User
						</button>
					</div>

					{/* Tab panels */}
					<div
						role="tabpanel"
						aria-labelledby={activeTab}
						style={{ marginTop: 12, padding: 12, border: '1px solid #e5e7eb', borderRadius: 8, background: '#fff' }}
					>
						{activeTab === 'workbook' ? (
							<div>
								<button
									type="button"
									onClick={() => setWorkbookModalOpen(true)}
									title="Open workbook inspector"
									aria-label="Open workbook inspector"
									style={{
										margin: '0 0 8px 0',
										padding: '4px 8px',
										borderRadius: 6,
										background: '#f3f4f6',
										border: '1px solid #e6e7eb',
										cursor: 'pointer',
										display: 'inline-flex',
										alignItems: 'center',
										gap: 6,
										fontSize: 15,
										fontWeight: 700
									}}
								>
									<span>Workbook Settings</span>
									{/* small file icon */}
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
										<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
										<path d="M14 2v6h6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
									</svg>
								</button>
								{renderSettingsControls(defaultWorkbookSettings, workbookSettingsState, updateWorkbookSetting, 'workbook-')}
							</div>
						) : activeTab === 'user' ? (
							<div>
								{/* User Information from SSO Token */}
								<UserInfoDisplay accessToken={accessToken} />

								{/* Power Automate License Section - Currently disabled due to Graph API requirements */}
								{/* Uncomment this section once you implement backend OBO flow */}
								{/*
								<div style={{ marginBottom: 16 }}>
									<h3 style={{ margin: '0 0 8px 0', fontSize: 14, fontWeight: 600, color: '#374151' }}>Power Automate License</h3>
									<LicenseChecker accessToken={accessToken} />
								</div>
								*/}
							</div>
						) : null}
					</div>
				</div>
			</div>

			{/* Render the modals via the new component */}
			<SettingsModal
				// array modal
				modalOpen={modalOpen}
				modalSetting={modalSetting}
				modalArray={modalArray}
				setModalArray={setModalArray}
				closeModal={closeModal}
				saveModal={saveModal}
				// provide current workbook columns so modal uses document columns instead of defaults
				workbookColumns={workbookSettingsState.columns}
				// master list headers to filter which columns are shown (null = show all)
				masterListHeaders={masterListHeaders}
				// set of ML headers that are new since last save (for "New" badge in Add picker)
				newMasterListHeaders={newMasterListHeaders}
				// set of ML headers whose data columns are entirely blank
				blankColumns={blankColumns}
			/>
			{/* pass the document key constant so the modal can read the workbook-specific mapping */}
			<WorkbookSettingsModal isOpen={workbookModalOpen} onClose={() => setWorkbookModalOpen(false)} docKey={DOC_KEY} />

			{/* Power Automate configuration modal */}
			<PowerAutomateConfigModal isOpen={powerAutomateModalOpen} onClose={() => setPowerAutomateModalOpen(false)} />
		</div>
	);
};

export default Settings;
