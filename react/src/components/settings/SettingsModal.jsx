import React from 'react';
import { X } from 'lucide-react';
import DeleteConfirmModal from './DeleteConfirmModal'; // <-- added

const SettingsModal = ({
	// array modal props
	modalOpen,
	modalSetting,
	modalArray = [],
	setModalArray = () => {},
	closeModal = () => {},
	saveModal = () => {},
	// new: prefer workbook columns (from saved workbook settings) over default choices
	workbookColumns = [],
	// master list headers to filter visible columns (null = show all)
	masterListHeaders = null
}) => {
	return (
		<>
			{/* Only show the editableArray modal */}
			{modalOpen && modalSetting?.type === 'array' ? (
				<div
					onClick={closeModal}
					style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}
					role="dialog"
					aria-modal="true"
				>
					<div onClick={e => e.stopPropagation()} style={{ width: 'min(820px, 96%)', maxHeight: '86vh', overflow: 'auto', background: '#fff', borderRadius: 8, padding: 16, boxSizing: 'border-box' }}>
						<div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
							<h3 style={{ margin: 0 }}>{modalSetting?.label || 'Edit items'}</h3>
							<button
								onClick={closeModal}
								style={{ background: 'transparent', border: 'none', cursor: 'pointer', padding: 6, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
								aria-label="Close"
							>
								<X size={16} />
							</button>
						</div>

						{/* editableArray content */}
						<EditableArrayInner
							modalSetting={modalSetting}
							modalArray={modalArray}
							setModalArray={setModalArray}
							closeModal={closeModal}
							saveModal={saveModal}
							// pass workbook columns so the modal uses the workbook's columns array when available
							workbookColumns={workbookColumns}
							masterListHeaders={masterListHeaders}
						/>
					</div>
				</div>
			) : null}
		</>
	);
};

export default SettingsModal;

// Replace the existing EditableArrayInner component with this updated version
const EditableArrayInner = ({ modalSetting, modalArray = [], setModalArray, closeModal, saveModal, workbookColumns = [], masterListHeaders = null }) => {
	const [selectedIdx, setSelectedIdx] = React.useState(0);
	const [viewMode, setViewMode] = React.useState('choices'); // 'choices' | 'options' | 'add'
	const [editableMap, setEditableMap] = React.useState({});
	const [hoverIdx, setHoverIdx] = React.useState(null);
	const [orderMap, setOrderMap] = React.useState({});
	const [orderList, setOrderList] = React.useState([]);
	// editing state for the "Options for:" title
	const [editingTitle, setEditingTitle] = React.useState(false);
	const [titleInput, setTitleInput] = React.useState('');
	// dirty tracking: initial snapshot + current dirty flag
	const initialSnapshotRef = React.useRef(null);
	const [isDirty, setIsDirty] = React.useState(false);
	// state to control delete confirmation modal
	const [deleteConfirmOpen, setDeleteConfirmOpen] = React.useState(false);
	// preview state
	const [previewBusy, setPreviewBusy] = React.useState(false);

	// build quick lookup of workbookColumns by key (name/label)
	const workbookLookup = React.useMemo(() => {
		const l = {};
		if (Array.isArray(workbookColumns)) {
			workbookColumns.forEach(entry => {
				const key = String(entry?.name ?? entry?.label ?? entry).trim();
				if (key) l[key] = entry;
			});
		}
		return l;
	}, [workbookColumns]);

	// helper for effect dependency without introducing new props
	function modalOpenOrSettingMarker() {
		return JSON.stringify({
			choices: (Array.isArray(workbookColumns) ? workbookColumns.length : 0),
			// include a small signature of workbook options so effect runs when options change
			workbookSignature: (Array.isArray(workbookColumns) ? workbookColumns.map(e => {
				const n = String(e?.name ?? e?.label ?? e).trim();
				const oLen = Array.isArray(e?.options) ? e.options.length : (e?.options && typeof e.options === 'object' ? Object.keys(e.options).length : 0);
				return `${n}:${oLen}`;
			}) : []).join('|'),
			options: modalSetting?.options?.length ?? 0,
			modalArrayLength: (modalArray || []).length
		});
	}

	// initialize editableMap, orderMap, orderList and selected index whenever the modal/setting/array changes
	React.useEffect(() => {
		const map = {};
		const order = {};
		let orderCounter = 0;
		// 1) seed map from modalArray (what was previously saved via Save)
		// Filter out columns that were previously hidden — they are now implicitly hidden by not being in the list
		if (Array.isArray(modalArray) && modalArray.length) {
			modalArray.forEach((entry) => {
				if (entry && entry.column) {
					const key = String(entry.column).trim();
					// entry.options may be an object with the per-column settings
					const fromOptions = (entry.options && typeof entry.options === 'object') ? { ...entry.options } : {};
					// also preserve top-level properties that may have been saved (alias/static/etc)
					const extra = {};
					['alias', 'static', 'format', 'label', 'name', 'identifier'].forEach(k => {
						if (entry[k] !== undefined) extra[k] = entry[k];
					});
					const merged = { ...fromOptions, ...extra };
					// Skip columns that were marked hidden — they belong in the "Add" bank now
					if (merged.hidden) return;
					delete merged.hidden; // clean up any residual hidden flag
					map[key] = merged;
					orderCounter++;
					order[key] = orderCounter;
				}
			});
		}

		// 2) determine the choices source (prefer workbookColumns when available)
		const choicesSource = (Array.isArray(workbookColumns) && workbookColumns.length) ? workbookColumns : (modalSetting?.choices || []);
		// helper to extract non-name/label properties from a workbook entry
		const seedFromWorkbookEntry = wbEntry => {
			const seed = {};
			if (!wbEntry || typeof wbEntry !== 'object') return seed;
			Object.keys(wbEntry).forEach(k => {
				if (k === 'name' || k === 'label') return;
				// copy everything else (alias, static, format, options, etc) but NOT hidden
				if (k === 'hidden') return;
				seed[k] = wbEntry[k];
			});
			return seed;
		};

		choicesSource.forEach((choice, i) => {
			const key = String(choice.name ?? choice.label ?? choice).trim();
			// Skip hidden defaults — they go to the "Add" bank
			if (choice.hidden) return;
			if (!map[key]) {
				// prefer seeding from the current workbook mapping if available
				const wbEntry = workbookLookup[key];
				if (wbEntry) {
					// skip if the workbook entry is hidden
					if (wbEntry.hidden) return;
					// if workbook provided an .options object/structure use it, otherwise seed from other properties
					if (wbEntry.options && typeof wbEntry.options === 'object' && !Array.isArray(wbEntry.options)) {
						map[key] = { ...(wbEntry.options) };
					} else {
						map[key] = seedFromWorkbookEntry(wbEntry);
					}
				} else {
					map[key] = {};
				}
			}
			if (!order[key]) {
				orderCounter++;
				order[key] = orderCounter;
			}
		});

		setEditableMap(map);
		setOrderMap(order);
		const orderedKeys = Object.keys(order).sort((a, b) => (order[a] || 0) - (order[b] || 0));
		setOrderList(orderedKeys);
		// reset title input when selection or order changes
		setTitleInput(orderedKeys[0] || '');
		setSelectedIdx(prev => (orderedKeys[prev] ? prev : 0));
		setViewMode('choices');

		// capture initial snapshot for dirty tracking whenever modal/setting opens or array changes
		try {
			initialSnapshotRef.current = JSON.stringify({ editableMap: map, orderList: orderedKeys });
			setIsDirty(false);
		} catch (e) {
			initialSnapshotRef.current = null;
			setIsDirty(false);
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [modalOpenOrSettingMarker()]);

	// update dirty state whenever editableMap or orderList change
	React.useEffect(() => {
		try {
			const cur = JSON.stringify({ editableMap, orderList });
			setIsDirty(initialSnapshotRef.current ? cur !== initialSnapshotRef.current : false);
		} catch (e) {
			setIsDirty(false);
		}
	}, [editableMap, orderList]);

	if (!modalSetting) return null;

	// prefer workbookColumns when available for the choice list shown in the UI
	const choices = (Array.isArray(workbookColumns) && workbookColumns.length) ? workbookColumns : (modalSetting.choices || []);
	// prefer options descriptors from the workbook entry for the currently selected key (if provided) else fall back to modalSetting.options
	const currentKey = orderList[selectedIdx] || '';
	const workbookEntryForCurrent = workbookLookup[currentKey];
	let options = Array.isArray(modalSetting?.options) ? modalSetting.options : [];
	if (workbookEntryForCurrent && Array.isArray(workbookEntryForCurrent.options) && workbookEntryForCurrent.options.length) {
		options = workbookEntryForCurrent.options;
	}

	const getChoiceByKey = key => {
		return (choices.find(c => {
			const k = String(c?.name ?? c?.label ?? c).trim();
			return k === key;
		}) || null);
	};

	const currentChoice = getChoiceByKey(currentKey);

	// helper: rename the current key (avoid duplicates)
	const renameCurrentKey = newLabel => {
		const newKey = String(newLabel || '').trim();
		if (!newKey || newKey === currentKey) return;
		if (orderList.includes(newKey)) {
			// duplicate - ignore rename
			return;
		}
		const next = [...orderList];
		next[selectedIdx] = newKey;
		const newOrderMap = {};
		next.forEach((k, i) => { newOrderMap[k] = i + 1; });
		setOrderList(next);
		setOrderMap(newOrderMap);
		setEditableMap(prev => {
			const copy = { ...prev };
			copy[newKey] = copy[currentKey] || {};
			delete copy[currentKey];
			return copy;
		});
		// keep selection on the renamed item
		setSelectedIdx(selectedIdx);
	};

	// when entering edit mode, prepare input
	React.useEffect(() => {
		if (editingTitle) {
			setTitleInput(currentKey || '');
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [editingTitle]);

	const updateOption = (optionName, value) => {
		setEditableMap(prev => {
			const copy = { ...prev };
			copy[currentKey] = { ...(copy[currentKey] || {}) };
			if (value === '' || value == null) {
				delete copy[currentKey][optionName];
			} else {
				copy[currentKey][optionName] = value;
			}
			return copy;
		});
	};

	const onSave = () => {
		const out = (orderList.length ? orderList : Object.keys(editableMap)).map(col => ({
			column: col,
			options: { ...(editableMap[col] || {}) }
		}));
		// update parent array state
		try {
			setModalArray(out);
		} catch (e) {
			// ignore
		}
		if (typeof saveModal === 'function') {
			try {
				saveModal(out);
			} catch (e) {
				// swallow errors, fallback already setModalArray
			}
		}

		// Do NOT close the modal. Instead go back to the choices list.
		setViewMode('choices');

		// update snapshot so Save becomes disabled until further edits
		try {
			initialSnapshotRef.current = JSON.stringify({ editableMap, orderList });
			setIsDirty(false);
		} catch (e) {
			// ignore
			setIsDirty(false);
		}
	};

	// OLD deleteCurrent replaced: open confirm modal instead of immediate confirm
	const deleteCurrent = () => {
		if (!currentKey) return;
		// open confirm modal — actual deletion is done in performDeleteCurrent
		setDeleteConfirmOpen(true);
	};

	// perform actual deletion when user confirms
	const performDeleteCurrent = () => {
		const keyToDelete = currentKey;
		if (!keyToDelete) {
			setDeleteConfirmOpen(false);
			return;
		}

		// Build new editable map & order list locally (synchronously) so we can persist immediately
		const newEditableMap = { ...(editableMap || {}) };
		delete newEditableMap[keyToDelete];

		const nextOrderList = orderList.filter(k => k !== keyToDelete);
		const newOrderMap = {};
		nextOrderList.forEach((k, i) => { newOrderMap[k] = i + 1; });

		// Update local UI state
		setEditableMap(newEditableMap);
		setOrderList(nextOrderList);
		setOrderMap(newOrderMap);
		const newIdx = Math.max(0, Math.min(nextOrderList.length - 1, selectedIdx));
		setSelectedIdx(newIdx);
		setTitleInput(nextOrderList[newIdx] || '');
		if (nextOrderList.length === 0) setViewMode('choices');

		// Build the array shape expected by the parent and persist immediately.
		const out = (nextOrderList.length ? nextOrderList : Object.keys(newEditableMap || {})).map(col => ({
			column: col,
			options: { ...(newEditableMap[col] || {}) }
		}));

		// Persist to parent if available; do NOT close the modal.
		if (typeof saveModal === 'function') {
			try {
				saveModal(out);
			} catch (e) {
				// fallback: update parent's modalArray locally
				try { setModalArray(out); } catch (_) {}
			}
		} else {
			// fallback local behavior
			setModalArray(out);
		}

		// update snapshot so Save becomes disabled until further edits
		try {
			initialSnapshotRef.current = JSON.stringify({ editableMap: newEditableMap, orderList: nextOrderList });
			setIsDirty(false);
		} catch (e) {
			setIsDirty(false);
		}

		// close the confirm dialog but keep the main modal open and show the list
		setDeleteConfirmOpen(false);
		setViewMode('choices');
	};

	// detect which columns in the visible list are missing from the current Master List
	const missingColumns = React.useMemo(() => {
		const missing = new Set();
		if (!Array.isArray(masterListHeaders) || masterListHeaders.length === 0) return missing;
		const stripStr = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
		const mlStripped = masterListHeaders.map(h => stripStr(h));
		orderList.forEach(key => {
			const keyStripped = stripStr(key);
			// check column name itself
			if (mlStripped.includes(keyStripped)) return;
			// check aliases from workbook entry or editableMap
			const wbEntry = workbookLookup[key];
			const aliases = [];
			if (wbEntry && Array.isArray(wbEntry.alias)) aliases.push(...wbEntry.alias);
			const mapEntry = editableMap[key];
			if (mapEntry && Array.isArray(mapEntry.alias)) aliases.push(...mapEntry.alias);
			if (mapEntry && typeof mapEntry.alias === 'string' && mapEntry.alias) {
				aliases.push(...mapEntry.alias.split(',').map(a => a.trim()).filter(Boolean));
			}
			const found = aliases.some(a => mlStripped.includes(stripStr(a)));
			if (!found) missing.add(key);
		});
		return missing;
	}, [orderList, masterListHeaders, workbookLookup, editableMap]);

	// move an item up/down within orderList
	const moveItem = (idx, direction) => {
		const targetIdx = idx + direction;
		if (targetIdx < 0 || targetIdx >= orderList.length) return;
		const next = [...orderList];
		[next[idx], next[targetIdx]] = [next[targetIdx], next[idx]];
		const newOrder = {};
		next.forEach((k, i) => { newOrder[k] = i + 1; });
		setOrderList(next);
		setOrderMap(newOrder);
	};

	// add a column from the master list (or blank)
	const addColumn = (name) => {
		if (orderList.includes(name)) return;
		const next = [...orderList, name];
		const newOrderMap = {};
		next.forEach((k, i) => { newOrderMap[k] = i + 1; });
		setOrderList(next);
		setOrderMap(newOrderMap);
		// seed from workbook column definition if available
		const wbEntry = workbookLookup[name];
		const seed = {};
		if (wbEntry && typeof wbEntry === 'object') {
			Object.keys(wbEntry).forEach(k => {
				if (k !== 'name' && k !== 'label') seed[k] = wbEntry[k];
			});
		}
		// new columns added from picker are visible by default
		delete seed.hidden;
		setEditableMap(prev => ({ ...(prev || {}), [name]: seed }));
		setIsDirty(true);
	};

	// compute master list columns not already configured (for the add picker)
	const availableToAdd = React.useMemo(() => {
		if (!Array.isArray(masterListHeaders)) return [];
		const existing = new Set(orderList.map(k => k.toLowerCase()));
		// also check aliases
		orderList.forEach(k => {
			const wbEntry = workbookLookup[k];
			if (wbEntry && Array.isArray(wbEntry.alias)) {
				wbEntry.alias.forEach(a => existing.add(String(a).trim().toLowerCase()));
			}
		});
		return masterListHeaders.filter(h => h && !existing.has(h.toLowerCase()));
	}, [masterListHeaders, orderList, workbookLookup]);

	// preview: create a headers-only LDA sheet
	const createPreview = async () => {
		if (previewBusy) return;
		setPreviewBusy(true);
		try {
			// first save current state
			onSave();
			if (typeof window !== 'undefined' && window.Excel && Excel.run) {
				await Excel.run(async (context) => {
					const sheets = context.workbook.worksheets;
					sheets.load('items/name');
					await context.sync();

					const today = new Date();
					const dateStr = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
					let sheetName = `LDA Preview ${dateStr}`;
					let counter = 2;
					const existingNames = sheets.items.map(s => s.name);
					while (existingNames.includes(sheetName)) {
						sheetName = `LDA Preview ${dateStr} (${counter++})`;
					}

					const newSheet = sheets.add(sheetName);
					// write visible column headers (skip missing columns)
					const previewHeaders = orderList.filter(key => !missingColumns.has(key));

					if (previewHeaders.length > 0) {
						const headerRange = newSheet.getRangeByIndexes(0, 0, 1, previewHeaders.length);
						headerRange.values = [previewHeaders];
						headerRange.format.font.bold = true;
						headerRange.format.autofitColumns();
					}

					newSheet.activate();
					await context.sync();
				});
			}
		} catch (err) {
			console.warn('Failed to create preview sheet', err);
		} finally {
			setPreviewBusy(false);
		}
	};

	// Render: add view (pick from master list)
	if (viewMode === 'add') {
		return (
			<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 8, background: '#fafafa', maxHeight: '56vh', overflowY: 'auto' }}>
				<div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
					<button
						onClick={() => setViewMode('choices')}
						style={{ padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb', background: '#f8fafc', cursor: 'pointer' }}
					>
						Back
					</button>
					<div style={{ fontWeight: 600, fontSize: 14 }}>Add columns from Master List</div>
				</div>
				{availableToAdd.length === 0 ? (
					<div style={{ color: '#6b7280', fontSize: 13 }}>All Master List columns are already configured.</div>
				) : (
					availableToAdd.map(header => (
						<button
							key={header}
							type="button"
							onClick={() => { addColumn(header); }}
							style={{
								display: 'flex',
								alignItems: 'center',
								width: '100%',
								padding: '8px 10px',
								marginBottom: 4,
								borderRadius: 6,
								border: '1px solid #e6e7eb',
								background: '#fff',
								cursor: 'pointer',
								textAlign: 'left',
								gap: 8,
								fontSize: 14,
								transition: 'background-color 120ms ease'
							}}
							onMouseEnter={e => { e.currentTarget.style.background = '#eef2ff'; }}
							onMouseLeave={e => { e.currentTarget.style.background = '#fff'; }}
						>
							<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0, color: '#4f46e5' }}>
								<path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
							</svg>
							{header}
						</button>
					))
				)}
			</div>
		);
	}

	// Render: choices view (visible columns with missing detection)
	if (viewMode === 'choices') {
		return (
			<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 8, background: '#fafafa', maxHeight: '56vh', overflowY: 'auto', position: 'relative' }}>
				{/* header: Visible on LDA + add/preview buttons */}
				<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
					<div style={{ fontWeight: 600, display: 'flex', alignItems: 'center', gap: 8, fontSize: 13 }}>Visible on LDA</div>
					<div style={{ display: 'flex', gap: 6 }}>
						<button
							type="button"
							onClick={createPreview}
							disabled={previewBusy || orderList.length === 0}
							title="Create a preview LDA sheet with just headers"
							style={{
								padding: '6px 8px',
								borderRadius: 6,
								background: previewBusy ? '#f3f4f6' : '#f0fdf4',
								border: '1px solid rgba(22,163,74,0.12)',
								cursor: (previewBusy || orderList.length === 0) ? 'not-allowed' : 'pointer',
								display: 'inline-flex',
								alignItems: 'center',
								gap: 6,
								fontSize: 13,
								color: '#15803d'
							}}
						>
							<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
								<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
								<circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="1.4"/>
							</svg>
							{previewBusy ? 'Creating...' : 'Preview'}
						</button>
						<button
							type="button"
							onClick={() => {
								if (Array.isArray(masterListHeaders) && masterListHeaders.length > 0) {
									setViewMode('add');
								} else {
									// no master list, add blank column
									const base = 'New Column';
									let counter = 1;
									let key = base;
									while (orderList.includes(key)) { counter++; key = `${base} ${counter}`; }
									addColumn(key);
								}
							}}
							title="Add column"
							aria-label="Add column"
							style={{
								padding: '6px 8px',
								borderRadius: 6,
								background: '#eef2ff',
								border: '1px solid rgba(79,70,229,0.12)',
								cursor: 'pointer',
								display: 'inline-flex',
								alignItems: 'center',
								gap: 6,
								fontSize: 13
							}}
						>
							<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
								<path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
							</svg>
							Add
						</button>
					</div>
				</div>

				{orderList.length === 0 && (
					<div style={{ color: '#6b7280', fontSize: 13, marginBottom: 8 }}>No visible columns. Click Add to include columns from the Master List.</div>
				)}

				{orderList.map((key, idx) => {
					const choiceObj = getChoiceByKey(key);
					const label = String(choiceObj?.name ?? choiceObj?.label ?? key).trim();
					const num = idx + 1;
					const hoverKey = 'col-' + idx;
					const isMissing = missingColumns.has(key);

					return (
						<div
							key={key + idx}
							onMouseEnter={() => setHoverIdx(hoverKey)}
							onMouseLeave={() => setHoverIdx(null)}
							style={{
								display: 'flex',
								alignItems: 'center',
								width: '100%',
								padding: '6px 8px',
								marginBottom: 4,
								borderRadius: 6,
								background: isMissing
									? '#f9fafb'
									: (hoverIdx === hoverKey) ? '#eef2ff' : 'transparent',
								border: isMissing
									? '1px solid #e5e7eb'
									: (hoverIdx === hoverKey) ? '1px solid rgba(79,70,229,0.12)' : '1px solid transparent',
								transition: 'background-color 120ms ease',
								gap: 4,
								opacity: isMissing ? 0.5 : 1
							}}
						>
							{/* position badge */}
							<div
								aria-hidden="true"
								style={{
									width: 28,
									height: 28,
									display: 'inline-flex',
									alignItems: 'center',
									justifyContent: 'center',
									borderRadius: 6,
									background: isMissing ? '#f3f4f6' : (hoverIdx === hoverKey) ? '#eef2ff' : '#f3f4f6',
									color: isMissing ? '#9ca3af' : '#374151',
									fontWeight: 700,
									flexShrink: 0,
									userSelect: 'none',
									fontSize: 13
								}}
								title={isMissing ? `Position ${num} (missing from Master List)` : `Position ${num}`}
							>
								{isMissing ? (
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
										<circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.6"/>
										<line x1="12" y1="8" x2="12" y2="12" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
										<circle cx="12" cy="16" r="0.5" fill="currentColor" stroke="currentColor" strokeWidth="1"/>
									</svg>
								) : num}
							</div>

							{/* label - clickable to open options */}
							<button
								onClick={() => { if (!isMissing) { setSelectedIdx(idx); setViewMode('options'); } }}
								style={{
									flex: 1,
									background: 'none',
									border: 'none',
									textAlign: 'left',
									cursor: isMissing ? 'default' : 'pointer',
									padding: '4px 0',
									overflow: 'hidden',
									textOverflow: 'ellipsis',
									whiteSpace: 'nowrap',
									fontSize: 14,
									color: isMissing ? '#9ca3af' : 'inherit'
								}}
							>
								{label}
							</button>

							{/* missing badge */}
							{isMissing && (
								<span
									title="This column was not found in the Master List and will be skipped"
									style={{
										padding: '2px 6px',
										borderRadius: 4,
										background: '#fef3c7',
										color: '#92400e',
										fontSize: 11,
										fontWeight: 600,
										flexShrink: 0,
										lineHeight: '16px'
									}}
								>
									Missing
								</span>
							)}

							{/* up/down arrow buttons */}
							{!isMissing && (
								<div style={{ display: 'flex', flexDirection: 'column', gap: 1, flexShrink: 0 }}>
									<button
										type="button"
										onClick={(e) => { e.stopPropagation(); moveItem(idx, -1); }}
										disabled={idx === 0}
										aria-label={`Move ${label} up`}
										title="Move up"
										style={{
											padding: 0,
											width: 22,
											height: 16,
											display: 'inline-flex',
											alignItems: 'center',
											justifyContent: 'center',
											border: '1px solid #e6e7eb',
											borderRadius: '4px 4px 0 0',
											background: idx === 0 ? '#f9fafb' : '#f3f4f6',
											cursor: idx === 0 ? 'not-allowed' : 'pointer',
											color: idx === 0 ? '#d1d5db' : '#374151',
											transition: 'background-color 120ms ease'
										}}
									>
										<svg width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
											<path d="M18 15l-6-6-6 6" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
										</svg>
									</button>
									<button
										type="button"
										onClick={(e) => { e.stopPropagation(); moveItem(idx, 1); }}
										disabled={idx === orderList.length - 1}
										aria-label={`Move ${label} down`}
										title="Move down"
										style={{
											padding: 0,
											width: 22,
											height: 16,
											display: 'inline-flex',
											alignItems: 'center',
											justifyContent: 'center',
											border: '1px solid #e6e7eb',
											borderRadius: '0 0 4px 4px',
											background: idx === orderList.length - 1 ? '#f9fafb' : '#f3f4f6',
											cursor: idx === orderList.length - 1 ? 'not-allowed' : 'pointer',
											color: idx === orderList.length - 1 ? '#d1d5db' : '#374151',
											transition: 'background-color 120ms ease'
										}}
									>
										<svg width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
											<path d="M6 9l6 6 6-6" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
										</svg>
									</button>
								</div>
							)}
						</div>
					);
				})}
			</div>
		);
	}

	// options view
	return (
		<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 12, background: '#fff', maxHeight: '56vh', overflowY: 'auto' }}>
			{/* Header: show back button to return to choices-only view */}
			<div style={{ marginBottom: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
				<div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
					<button
						onClick={() => setViewMode('choices')}
						style={{ padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb', background: '#f8fafc', cursor: 'pointer' }}
						aria-label="Back to columns"
					>
						Back
					</button>
					<div style={{ fontWeight: 600, display: 'flex', alignItems: 'center', gap: 8 }}>
						<span>Options for:</span>
						{editingTitle ? (
							<input
								autoFocus
								value={titleInput}
								onChange={e => setTitleInput(e.target.value)}
								onBlur={() => {
									renameCurrentKey(titleInput);
									setEditingTitle(false);
								}}
								onKeyDown={e => {
									if (e.key === 'Enter') {
										renameCurrentKey(titleInput);
										setEditingTitle(false);
									} else if (e.key === 'Escape') {
										setEditingTitle(false);
									}
								}}
								style={{ padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb', minWidth: 160 }}
								aria-label="Rename column"
							/>
						) : (
							<span
								onClick={() => { setEditingTitle(true); }}
								style={{ fontWeight: 700, cursor: 'pointer', userSelect: 'none' }}
								title="Click to rename"
								role="button"
								tabIndex={0}
								onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { setEditingTitle(true); } }}
							>
								{currentKey}
							</span>
						)}
					</div>
				</div>
				<div style={{ fontSize: 13, color: '#6b7280' }}>{Object.keys(editableMap[currentKey] || {}).length} set</div>
			</div>

			<div style={{ display: 'grid', gap: 8 }}>
				{options.length === 0 && <div style={{ color: '#6b7280', fontSize: 13 }}>No options available</div>}
				{options.map(opt => {
					// use opt.option / opt.name as the internal key, but use opt.label (if provided) for display
					const optKey = opt.option || opt.name || String(opt);
					const optLabel = opt.label || optKey;
					const values = Array.isArray(opt.values) ? opt.values : (Array.isArray(opt.type) ? opt.type : []);
					const value = (editableMap[currentKey] && editableMap[currentKey][optKey]) || '';
					return (
						<div key={optKey} style={{ display: 'flex', gap: 8, alignItems: 'center', justifyContent: 'space-between' }}>
							<div style={{ minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{optLabel}</div>

							{values && values.length > 0 ? (
								<select
									value={value}
									onChange={e => updateOption(optKey, e.target.value)}
									style={{ padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb', minWidth: 160 }}
									aria-label={`Select ${optLabel} for ${currentKey}`}
								>
									<option value=''>None</option>
									{values.map(v => (
										<option key={v} value={v}>{v}</option>
									))}
								</select>
							) : opt.type === 'boolean' ? (
								(() => {
									const checked = value === true || value === 'Yes' || value === 'On' || String(value).toLowerCase() === 'true' || String(value).toLowerCase() === 'yes' || String(value).toLowerCase() === 'on';
									return (
										<label style={{ display: 'inline-flex', alignItems: 'center', gap: 1, cursor: 'pointer' }}>
											<span style={{ position: 'relative', width: 44, height: 24, display: 'inline-block' }}>
												<input
													type="checkbox"
													checked={checked}
													onChange={e => updateOption(optKey, e.target.checked)}
													aria-label={optLabel}
													style={{ position: 'absolute', opacity: 0, width: 0, height: 0 }}
												/>
												<span style={{
													position: 'absolute',
													inset: 0,
													borderRadius: 9999,
													background: checked ? '#4f46e5' : '#e5e7eb',
													transition: 'background-color 160ms linear',
													padding: 2,
													boxSizing: 'border-box'
												}} />
												<span style={{
													position: 'absolute',
													top: 2,
													left: checked ? 22 : 2,
													width: 20,
													height: 20,
													borderRadius: '50%',
													background: '#fff',
													boxShadow: '0 1px 2px rgba(0,0,0,0.2)',
													transition: 'left 160ms linear',
												}} />
											</span>
											<span style={{ fontSize: 13 }}>{checked ? 'On' : 'Off'}</span>
										</label>
									);
								})()
							) : (
								<input
									type="text"
									value={value}
									onChange={e => updateOption(optKey, e.target.value)}
									style={{ padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb', minWidth: 160 }}
									aria-label={`Enter ${optLabel} for ${currentKey}`}
									placeholder="Enter value"
								/>
							)}
						</div>
					);
				})}
			</div>

			{/* Actions */}
			<div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 12 }}>
				{/* Delete current column (left) */}
				<button
					onClick={deleteCurrent}
					disabled={!currentKey}
					aria-disabled={!currentKey}
					title={currentKey ? `Delete ${currentKey}` : 'No column selected'}
					style={{
						padding: '8px 10px',
						borderRadius: 6,
						background: '#fff',
						color: '#ef4444',
						border: '1px solid rgba(239,68,68,0.12)',
						cursor: currentKey ? 'pointer' : 'not-allowed',
						opacity: currentKey ? 1 : 0.6
					}}
				>
					Delete
				</button>

				<button
					onClick={onSave}
					disabled={!isDirty}
					aria-disabled={!isDirty}
					style={{
						padding: '8px 10px',
						borderRadius: 6,
						background: isDirty ? '#4f46e5' : '#9ca3af',
						color: '#fff',
						border: 'none',
						cursor: isDirty ? 'pointer' : 'not-allowed',
						opacity: isDirty ? 1 : 0.9
					}}
					title={isDirty ? 'Save changes' : 'No changes to save'}
				>
					Save
				</button>
			</div>

			{/* Delete confirmation modal */}
			<DeleteConfirmModal
				isOpen={deleteConfirmOpen}
				title="Delete column"
				message={`Delete column "${currentKey}"? This cannot be undone.`}
				confirmLabel="Delete"
				onConfirm={performDeleteCurrent}
				onCancel={() => setDeleteConfirmOpen(false)}
			/>
		</div>
	);
};
