import React, { useState } from 'react';
import { Info } from 'lucide-react'; // removed X + Plus imports (moved to SettingsModal)
import '../studentView/Styling/StudentView.css'; // add StudentView tab styles
import { defaultUserSettings, defaultWorkbookSettings, sectionIcons } from './DefaultSettings'; // added: import defaults
import SettingsModal from './SettingsModal'; // new: modal component

const Settings = () => {
	// placeholder SVG avatar
	const svg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'>
		<rect fill='%23e5e7eb' width='100' height='100' rx='16'/>
		<circle fill='%239ca3af' cx='50' cy='36' r='18'/>
		<path fill='%239ca3af' d='M20 84c0-14 12-26 30-26s30 12 30 26' />
		</svg>`;
	const avatarSrc = `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;

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

	const openArrayModal = (setting, currentValue, updater) => {
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

	const saveModal = () => {
		if (modalSetting && typeof modalUpdater === 'function') {
			// ensure we send a clean array of objects: { name, alias: string[], static: boolean }
			const cleaned = modalArray.map(it => ({
				name: String(it?.name ?? '').trim(),
				alias: Array.isArray(it?.alias) ? it.alias.map(a => String(a).trim()).filter(Boolean) : [],
				static: !!it?.static
			}));
			modalUpdater(modalSetting.id, cleaned);
		}
		closeModal();
	};

	// helper to update a workbook setting
	const updateWorkbookSetting = (id, value) => {
		setWorkbookSettingsState(prev => ({ ...prev, [id]: value }));
	};

	// helper to update a setting
	const updateSetting = (id, value) => {
		setUserSettingsState(prev => ({ ...prev, [id]: value }));
	};

	// --- NEW: selections modal state & helpers ---
	const [selectionsModalOpen, setSelectionsModalOpen] = useState(false);
	const [selectionsModalSetting, setSelectionsModalSetting] = useState(null);
	const [selectionsAvailable, setSelectionsAvailable] = useState([]); // items not chosen
	const [selectionsChosen, setSelectionsChosen] = useState([]);       // items chosen
	const [selectionsFilter, setSelectionsFilter] = useState('');
	const [selectionsUpdater, setSelectionsUpdater] = useState(null);

	// helper to normalize an array of option items into { key, label, raw }
	const normalizeOptions = arr => {
		// Accept primitives or objects. Prefer common choice shapes:
		// { value, label }, { id, name }, { id, value, label }, { title, name }, etc.
		return (Array.isArray(arr) ? arr : []).map((it, i) => {
			if (typeof it === 'string') return { key: `${i}-${String(it)}`, label: it, raw: it };
			if (it && typeof it === 'object') {
				// pick friendly label from common fields
				const label = it.label ?? it.name ?? it.title ?? String(it.id ?? it.value ?? JSON.stringify(it));
				// prefer explicit value/id for identity
				const key = String(it.value ?? it.id ?? `${i}-${label}`);
				return { key, label, raw: it };
			}
			return { key: `${i}-item`, label: String(it), raw: it };
		});
	};

	const openSelectionsModal = (setting, currentValue, updater) => {
		// Use setting.choices if provided, otherwise fall back to setting.options
		const source = setting.choices ?? setting.options ?? [];
		// normalize all available options from the provided choices/options
		const all = normalizeOptions(source);
		// normalize current chosen value (could be array of strings/objects)
		const chosenNormalized = normalizeOptions(currentValue ?? []);
		// build chosen set using key equality
		const chosenKeys = new Set(chosenNormalized.map(c => c.key));
		// split available vs chosen ensuring no duplicates
		const available = all.filter(a => !chosenKeys.has(a.key));
		const chosen = all.filter(a => chosenKeys.has(a.key)).concat(
			// also include any chosen items not present in all (custom items)
			chosenNormalized.filter(c => !all.some(a => a.key === c.key))
		);
		setSelectionsAvailable(available);
		setSelectionsChosen(chosen);
		setSelectionsModalSetting(setting);
		setSelectionsUpdater(() => updater);
		setSelectionsFilter('');
		setSelectionsModalOpen(true);
	};

	const closeSelectionsModal = () => {
		setSelectionsModalOpen(false);
		setSelectionsModalSetting(null);
		setSelectionsAvailable([]);
		setSelectionsChosen([]);
		setSelectionsUpdater(null);
		setSelectionsFilter('');
	};

	const saveSelectionsModal = () => {
		if (selectionsModalSetting && typeof selectionsUpdater === 'function') {
			// send the raw values back to updater in the original shape
			const cleaned = selectionsChosen.map(i => i.raw);
			selectionsUpdater(selectionsModalSetting.id, cleaned);
		}
		closeSelectionsModal();
	};

	const moveToChosen = item => {
		setSelectionsAvailable(prev => prev.filter(i => i.key !== item.key));
		setSelectionsChosen(prev => [...prev, item]);
	};

	const moveToAvailable = item => {
		setSelectionsChosen(prev => prev.filter(i => i.key !== item.key));
		setSelectionsAvailable(prev => [...prev, item]);
	};
	// --- END NEW ---

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

							// ADD: show Configure button for editableArray
							if (setting.type === 'editableArray') {
								return (
									<button
										onClick={() => openArrayModal(setting, cur, updater)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#eef2ff', border: '1px solid #e0e7ff', cursor: 'pointer' }}
										aria-label={`Configure`}
									>
										Configure
									</button>
								);
							}

							// NEW: selections type shows a two-bank configure flow
							if (setting.type === 'selections') {
								return (
									<button
										onClick={() => openSelectionsModal(setting, cur, updater)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#eef2ff', border: '1px solid #e0e7ff', cursor: 'pointer' }}
										aria-label={`Configure`}
									>
										Configure
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

	// determine which image-type setting to use for the top avatar (if any)
	const imageSetting = defaultUserSettings.find(s => s.type === 'image');
	const avatarDisplaySrc = imageSetting
		? (userSettingsState[imageSetting.id] || imageSetting.defaultValue || avatarSrc)
		: avatarSrc;

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
					<h1 style={{ margin: 0, backgroundColor: '#f3f4f6', padding: '6px 8px', borderRadius: 6 }}>Welcome to Settings</h1>
					<p style={{ margin: '4px 0 12px 0' }}>
						This is a placeholder settings page. Configure your add-in options here.
					</p>

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
								<h2 style={{ margin: '0 0 8px 0', backgroundColor: '#f3f4f6', padding: '4px 8px', borderRadius: 6 }}>Workbook Settings</h2>
								{renderSettingsControls(defaultWorkbookSettings, workbookSettingsState, updateWorkbookSetting, 'workbook-')}
							</div>
						) : (
							<div>
								<h2 style={{ margin: '0 0 8px 0', backgroundColor: '#f3f4f6', padding: '4px 8px', borderRadius: 6 }}>User Settings</h2>

								{/* Render each default user setting as a label + button */}
								{renderSettingsControls(defaultUserSettings, userSettingsState, updateSetting)}

							</div>
						)}
					</div>
				</div>

				<img
					src={avatarDisplaySrc}
					alt="Profile placeholder"
					width="64"
					height="64"
					style={{
						position: 'absolute',
						top: 12,
						right: 12,
						width: 64,
						height: 64,
						borderRadius: '50%',
						objectFit: 'cover',       // ensure aspect ratio preserved and image is cropped to fill
						objectPosition: 'center', // center the crop (zoom)
						zIndex: 10,
					}}
				/>
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

				// selections modal
				selectionsModalOpen={selectionsModalOpen}
				selectionsModalSetting={selectionsModalSetting}
				selectionsAvailable={selectionsAvailable}
				selectionsChosen={selectionsChosen}
				selectionsFilter={selectionsFilter}
				setSelectionsFilter={setSelectionsFilter}
				closeSelectionsModal={closeSelectionsModal}
				saveSelectionsModal={saveSelectionsModal}
				moveToChosen={moveToChosen}
				moveToAvailable={moveToAvailable}
			/>
		</div>
	);
};

export default Settings;
