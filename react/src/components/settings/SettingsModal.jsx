import React from 'react';
import { X, GripVertical, Trash2 } from 'lucide-react';
import {
	DndContext,
	closestCenter,
	KeyboardSensor,
	PointerSensor,
	useSensor,
	useSensors,
} from '@dnd-kit/core';
import {
	arrayMove,
	SortableContext,
	sortableKeyboardCoordinates,
	useSortable,
	verticalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';

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
	masterListHeaders = null,
	// set of ML headers that are new since last save (for "New" badge)
	newMasterListHeaders = null,
	// set of ML headers whose data columns are entirely blank
	blankColumns = null
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
							saveModal={saveModal}
							workbookColumns={workbookColumns}
							masterListHeaders={masterListHeaders}
							newMasterListHeaders={newMasterListHeaders}
							blankColumns={blankColumns}
						/>
					</div>
				</div>
			) : null}
		</>
	);
};

export default SettingsModal;

// Sortable row used inside the visible-columns list
const SortableRow = ({
	id,
	idx,
	label,
	isMissing,
	isBlank,
	onRequestDelete,
}) => {
	const {
		attributes,
		listeners,
		setNodeRef,
		transform,
		transition,
		isDragging,
	} = useSortable({ id, disabled: isMissing });

	const [trashHover, setTrashHover] = React.useState(false);
	const [rowHover, setRowHover] = React.useState(false);

	const style = {
		transform: CSS.Transform.toString(transform),
		transition,
		display: 'flex',
		alignItems: 'center',
		width: '100%',
		padding: '6px 8px',
		marginBottom: 4,
		borderRadius: 6,
		background: isMissing
			? '#f9fafb'
			: (rowHover || isDragging) ? '#eef2ff' : 'transparent',
		border: isMissing
			? '1px solid #e5e7eb'
			: (rowHover || isDragging) ? '1px solid rgba(79,70,229,0.12)' : '1px solid transparent',
		gap: 4,
		opacity: isDragging ? 0.6 : (isMissing ? 0.5 : 1),
		boxShadow: isDragging ? '0 4px 12px rgba(0,0,0,0.12)' : 'none',
		zIndex: isDragging ? 1 : 'auto',
		userSelect: 'none',
		WebkitUserSelect: 'none',
	};

	const num = idx + 1;

	return (
		<div
			ref={setNodeRef}
			style={style}
			onMouseEnter={() => setRowHover(true)}
			onMouseLeave={() => setRowHover(false)}
		>
			{/* drag handle */}
			<button
				type="button"
				{...attributes}
				{...listeners}
				disabled={isMissing}
				aria-label={isMissing ? `${label} (missing, cannot reorder)` : `Drag to reorder ${label}`}
				title={isMissing ? 'Missing column cannot be reordered' : 'Drag to reorder'}
				style={{
					padding: 0,
					width: 22,
					height: 28,
					display: 'inline-flex',
					alignItems: 'center',
					justifyContent: 'center',
					border: 'none',
					background: 'transparent',
					cursor: isMissing ? 'not-allowed' : (isDragging ? 'grabbing' : 'grab'),
					color: isMissing ? '#d1d5db' : '#9ca3af',
					flexShrink: 0,
					touchAction: 'none',
				}}
			>
				<GripVertical size={16} />
			</button>

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
					background: isMissing ? '#f3f4f6' : (rowHover ? '#eef2ff' : '#f3f4f6'),
					color: isMissing ? '#9ca3af' : '#374151',
					fontWeight: 700,
					flexShrink: 0,
					userSelect: 'none',
					fontSize: 13,
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

			{/* label */}
			<div
				style={{
					flex: 1,
					padding: '4px 0',
					overflow: 'hidden',
					textOverflow: 'ellipsis',
					whiteSpace: 'nowrap',
					fontSize: 14,
					color: isMissing ? '#9ca3af' : 'inherit',
				}}
			>
				{label}
			</div>

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
						lineHeight: '16px',
					}}
				>
					Missing
				</span>
			)}

			{/* blank badge */}
			{!isMissing && isBlank && (
				<span
					title="This column exists in the Master List but contains no data values"
					style={{
						padding: '2px 6px',
						borderRadius: 4,
						background: '#e0e7ff',
						color: '#3730a3',
						fontSize: 11,
						fontWeight: 600,
						flexShrink: 0,
						lineHeight: '16px',
					}}
				>
					Blank
				</span>
			)}

			{/* trash icon delete button */}
			<button
				type="button"
				onClick={(e) => { e.stopPropagation(); onRequestDelete(id); }}
				onMouseEnter={() => setTrashHover(true)}
				onMouseLeave={() => setTrashHover(false)}
				aria-label={`Delete ${label}`}
				title={`Delete ${label}`}
				style={{
					padding: 0,
					width: 28,
					height: 28,
					display: 'inline-flex',
					alignItems: 'center',
					justifyContent: 'center',
					border: 'none',
					background: trashHover ? '#fee2e2' : 'transparent',
					borderRadius: 6,
					cursor: 'pointer',
					color: trashHover ? '#ef4444' : '#9ca3af',
					transition: 'background-color 120ms ease, color 120ms ease',
					flexShrink: 0,
				}}
			>
				<Trash2 size={15} />
			</button>
		</div>
	);
};

const EditableArrayInner = ({ modalSetting, modalArray = [], setModalArray, saveModal, workbookColumns = [], masterListHeaders = null, newMasterListHeaders = null, blankColumns = null }) => {
	const [viewMode, setViewMode] = React.useState('choices'); // 'choices' | 'add'
	const [editableMap, setEditableMap] = React.useState({});
	const [orderList, setOrderList] = React.useState([]);
	// dirty tracking: initial snapshot + current dirty flag
	const initialSnapshotRef = React.useRef(null);
	const [isDirty, setIsDirty] = React.useState(false);
	// preview state
	const [previewBusy, setPreviewBusy] = React.useState(false);
	// columns the user removed during this modal session (cleared on reopen)
	// Map<columnName, originalIndex> so re-adding restores the original position
	const [recentlyRemoved, setRecentlyRemoved] = React.useState(() => new Map());

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
			workbookSignature: (Array.isArray(workbookColumns) ? workbookColumns.map(e => {
				const n = String(e?.name ?? e?.label ?? e).trim();
				const oLen = Array.isArray(e?.options) ? e.options.length : (e?.options && typeof e.options === 'object' ? Object.keys(e.options).length : 0);
				return `${n}:${oLen}`;
			}) : []).join('|'),
			options: modalSetting?.options?.length ?? 0,
			modalArrayLength: (modalArray || []).length
		});
	}

	// initialize editableMap, orderList whenever the modal/setting/array changes
	React.useEffect(() => {
		const map = {};
		const order = [];
		let hasSavedEntries = false;
		// 1) seed map from modalArray (what was previously saved via Save)
		if (Array.isArray(modalArray) && modalArray.length) {
			modalArray.forEach((entry) => {
				const rawKey = entry?.column ?? entry?.name;
				if (entry && rawKey) {
					const key = String(rawKey).trim();
					const fromOptions = (entry.options && typeof entry.options === 'object') ? { ...entry.options } : {};
					const extra = {};
					['alias', 'static', 'format', 'label', 'name', 'identifier'].forEach(k => {
						if (entry[k] !== undefined) extra[k] = entry[k];
					});
					const merged = { ...fromOptions, ...extra };
					if (merged.hidden) return;
					delete merged.hidden;
					map[key] = merged;
					order.push(key);
					hasSavedEntries = true;
				}
			});
		}

		// 2) Only seed from choices source on FIRST-TIME setup (no saved entries).
		if (!hasSavedEntries) {
			const choicesSource = (Array.isArray(workbookColumns) && workbookColumns.length) ? workbookColumns : (modalSetting?.choices || []);
			const seedFromWorkbookEntry = wbEntry => {
				const seed = {};
				if (!wbEntry || typeof wbEntry !== 'object') return seed;
				Object.keys(wbEntry).forEach(k => {
					if (k === 'name' || k === 'label') return;
					if (k === 'hidden') return;
					seed[k] = wbEntry[k];
				});
				return seed;
			};

			choicesSource.forEach((choice) => {
				const key = String(choice.name ?? choice.label ?? choice).trim();
				if (choice.hidden) return;
				if (!map[key]) {
					const wbEntry = workbookLookup[key];
					if (wbEntry) {
						if (wbEntry.hidden) return;
						if (wbEntry.options && typeof wbEntry.options === 'object' && !Array.isArray(wbEntry.options)) {
							map[key] = { ...(wbEntry.options) };
						} else {
							map[key] = seedFromWorkbookEntry(wbEntry);
						}
					} else {
						map[key] = {};
					}
				}
				if (!order.includes(key)) order.push(key);
			});
		} else {
			// Even with saved entries, enrich map with workbook metadata
			Object.keys(map).forEach(key => {
				const wbEntry = workbookLookup[key];
				if (wbEntry && typeof wbEntry === 'object') {
					const existing = map[key];
					['alias', 'static', 'format', 'identifier'].forEach(k => {
						if (existing[k] === undefined && wbEntry[k] !== undefined) {
							existing[k] = wbEntry[k];
						}
					});
				}
			});
		}

		setEditableMap(map);
		setOrderList(order);
		setViewMode('choices');
		setRecentlyRemoved(new Map());

		// capture initial snapshot for dirty tracking whenever modal/setting opens or array changes
		try {
			initialSnapshotRef.current = JSON.stringify({ editableMap: map, orderList: order });
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

	// dnd-kit sensors
	const sensors = useSensors(
		useSensor(PointerSensor, { activationConstraint: { distance: 5 } }),
		useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
	);

	// detect which columns in the visible list are missing from the current Master List
	const missingColumns = React.useMemo(() => {
		const missing = new Set();
		if (!Array.isArray(masterListHeaders) || masterListHeaders.length === 0) return missing;
		const stripStr = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
		const mlStripped = masterListHeaders.map(h => stripStr(h));
		orderList.forEach(key => {
			const keyStripped = stripStr(key);
			if (mlStripped.includes(keyStripped)) return;
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

	// detect which columns in the visible list are blank
	const blankVisibleColumns = React.useMemo(() => {
		const blank = new Set();
		if (!blankColumns || blankColumns.size === 0) return blank;
		if (!Array.isArray(masterListHeaders) || masterListHeaders.length === 0) return blank;
		const stripStr = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
		const blankStripped = new Set();
		blankColumns.forEach(h => blankStripped.add(stripStr(h)));
		orderList.forEach(key => {
			if (missingColumns.has(key)) return;
			const keyStripped = stripStr(key);
			if (blankStripped.has(keyStripped)) { blank.add(key); return; }
			const wbEntry = workbookLookup[key];
			const aliases = [];
			if (wbEntry && Array.isArray(wbEntry.alias)) aliases.push(...wbEntry.alias);
			const mapEntry = editableMap[key];
			if (mapEntry && Array.isArray(mapEntry.alias)) aliases.push(...mapEntry.alias);
			if (mapEntry && typeof mapEntry.alias === 'string' && mapEntry.alias) {
				aliases.push(...mapEntry.alias.split(',').map(a => a.trim()).filter(Boolean));
			}
			const found = aliases.some(a => blankStripped.has(stripStr(a)));
			if (found) blank.add(key);
		});
		return blank;
	}, [orderList, blankColumns, masterListHeaders, missingColumns, workbookLookup, editableMap]);

	// drag-end handler: reorder the list
	const handleDragEnd = (event) => {
		const { active, over } = event;
		if (!over || active.id === over.id) return;
		setOrderList((items) => {
			const oldIndex = items.indexOf(active.id);
			const newIndex = items.indexOf(over.id);
			if (oldIndex === -1 || newIndex === -1) return items;
			return arrayMove(items, oldIndex, newIndex);
		});
	};

	// add a column from the master list (or blank)
	const addColumn = (name) => {
		if (orderList.includes(name)) return;
		const restoreIdx = recentlyRemoved.has(name) ? recentlyRemoved.get(name) : -1;
		setOrderList(prev => {
			if (restoreIdx < 0) return [...prev, name];
			const insertAt = Math.max(0, Math.min(prev.length, restoreIdx));
			const next = [...prev];
			next.splice(insertAt, 0, name);
			return next;
		});
		const wbEntry = workbookLookup[name];
		const seed = {};
		if (wbEntry && typeof wbEntry === 'object') {
			Object.keys(wbEntry).forEach(k => {
				if (k !== 'name' && k !== 'label') seed[k] = wbEntry[k];
			});
		}
		delete seed.hidden;
		setEditableMap(prev => ({ ...(prev || {}), [name]: seed }));
		// re-adding a recently-removed column clears its pill
		setRecentlyRemoved(prev => {
			if (!prev.has(name)) return prev;
			const next = new Map(prev);
			next.delete(name);
			return next;
		});
	};

	// delete a column locally (Save persists)
	const deleteColumn = (key) => {
		if (!key) return;
		const originalIdx = orderList.indexOf(key);
		setEditableMap(prev => {
			const copy = { ...(prev || {}) };
			delete copy[key];
			return copy;
		});
		setOrderList(prev => prev.filter(k => k !== key));
		setRecentlyRemoved(prev => {
			const next = new Map(prev);
			// keep the first-recorded original index so multiple delete/re-add cycles still snap back
			if (!next.has(key)) next.set(key, originalIdx);
			return next;
		});
	};

	const onSave = () => {
		const out = (orderList.length ? orderList : Object.keys(editableMap)).map(col => ({
			column: col,
			options: { ...(editableMap[col] || {}) }
		}));
		try {
			setModalArray(out);
		} catch (e) {
			// ignore
		}
		if (typeof saveModal === 'function') {
			try {
				saveModal(out);
			} catch (e) {
				// fallback already setModalArray
			}
		}
	};

	// compute columns not already configured (for the add picker)
	// includes master-list columns plus any column the user removed during this session
	const availableToAdd = React.useMemo(() => {
		const existing = new Set(orderList.map(k => k.toLowerCase()));
		orderList.forEach(k => {
			const wbEntry = workbookLookup[k];
			if (wbEntry && Array.isArray(wbEntry.alias)) {
				wbEntry.alias.forEach(a => existing.add(String(a).trim().toLowerCase()));
			}
		});
		const seen = new Set();
		const result = [];
		const push = (h) => {
			if (!h) return;
			const lower = String(h).toLowerCase();
			if (existing.has(lower) || seen.has(lower)) return;
			seen.add(lower);
			result.push(h);
		};
		recentlyRemoved.forEach((_, key) => push(key));
		if (Array.isArray(masterListHeaders)) masterListHeaders.forEach(push);
		return result;
	}, [masterListHeaders, orderList, workbookLookup, recentlyRemoved]);

	// preview: create a headers-only LDA sheet
	const createPreview = async () => {
		if (previewBusy) return;
		setPreviewBusy(true);
		try {
			if (typeof window !== 'undefined' && window.Excel && Excel.run) {
				await Excel.run(async (context) => {
					const sheets = context.workbook.worksheets;
					sheets.load('items/name');
					await context.sync();

					let sheetName = 'LDA Preview';
					let counter = 2;
					const existingNames = sheets.items.map(s => s.name);
					while (existingNames.includes(sheetName)) {
						sheetName = `LDA Preview (${counter++})`;
					}

					const newSheet = sheets.add(sheetName);
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

	const getChoiceByKey = key => {
		const choices = (Array.isArray(workbookColumns) && workbookColumns.length) ? workbookColumns : (modalSetting?.choices || []);
		return (choices.find(c => {
			const k = String(c?.name ?? c?.label ?? c).trim();
			return k === key;
		}) || null);
	};

	// Render: add view (pick from master list)
	if (viewMode === 'add') {
		const rank = (h) => {
			if (recentlyRemoved.has(h)) return 0;
			if (newMasterListHeaders && newMasterListHeaders.has(h)) return 1;
			return 2;
		};
		const sortedAvailable = [...availableToAdd].sort((a, b) => rank(a) - rank(b));
		const newCount = newMasterListHeaders ? sortedAvailable.filter(h => newMasterListHeaders.has(h) && !recentlyRemoved.has(h)).length : 0;
		const removedCount = sortedAvailable.filter(h => recentlyRemoved.has(h)).length;

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
					{removedCount > 0 && (
						<span style={{
							padding: '2px 8px',
							borderRadius: 10,
							background: '#fee2e2',
							color: '#b91c1c',
							fontSize: 11,
							fontWeight: 600,
							lineHeight: '16px'
						}}>
							{removedCount} recently removed
						</span>
					)}
					{newCount > 0 && (
						<span style={{
							padding: '2px 8px',
							borderRadius: 10,
							background: '#dbeafe',
							color: '#1d4ed8',
							fontSize: 11,
							fontWeight: 600,
							lineHeight: '16px'
						}}>
							{newCount} new
						</span>
					)}
				</div>
				{sortedAvailable.length === 0 ? (
					<div style={{ color: '#6b7280', fontSize: 13 }}>All Master List columns are already configured.</div>
				) : (
					sortedAvailable.map(header => {
						const isRemoved = recentlyRemoved.has(header);
						const isNew = !isRemoved && newMasterListHeaders && newMasterListHeaders.has(header);
						const restColor = isRemoved ? '#fef2f2' : (isNew ? '#eff6ff' : '#fff');
						const borderColor = isRemoved ? '1px solid rgba(239,68,68,0.25)' : (isNew ? '1px solid rgba(29,78,216,0.2)' : '1px solid #e6e7eb');
						return (
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
									border: borderColor,
									background: restColor,
									cursor: 'pointer',
									textAlign: 'left',
									gap: 8,
									fontSize: 14,
									transition: 'background-color 120ms ease'
								}}
								onMouseEnter={e => { e.currentTarget.style.background = isRemoved ? '#fee2e2' : '#eef2ff'; }}
								onMouseLeave={e => { e.currentTarget.style.background = restColor; }}
							>
								<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0, color: isRemoved ? '#ef4444' : '#4f46e5' }}>
									<path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
								</svg>
								<span style={{ flex: 1 }}>{header}</span>
								{isRemoved && (
									<span
										title="You removed this column in this session — click to add it back to its original position"
										style={{
											padding: '2px 6px',
											borderRadius: 4,
											background: '#fee2e2',
											color: '#b91c1c',
											fontSize: 11,
											fontWeight: 600,
											flexShrink: 0,
											lineHeight: '16px'
										}}
									>
										Recently Removed
									</span>
								)}
								{isNew && (
									<span
										title="This column was recently added to the Master List"
										style={{
											padding: '2px 6px',
											borderRadius: 4,
											background: '#dbeafe',
											color: '#1d4ed8',
											fontSize: 11,
											fontWeight: 600,
											flexShrink: 0,
											lineHeight: '16px'
										}}
									>
										New
									</span>
								)}
							</button>
						);
					})
				)}
			</div>
		);
	}

	// Render: choices view (visible columns with drag-and-drop reordering)
	return (
		<div>
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

				<DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
					<SortableContext items={orderList} strategy={verticalListSortingStrategy}>
						{orderList.map((key, idx) => {
							const choiceObj = getChoiceByKey(key);
							const label = String(choiceObj?.name ?? choiceObj?.label ?? key).trim();
							return (
								<SortableRow
									key={key}
									id={key}
									idx={idx}
									label={label}
									isMissing={missingColumns.has(key)}
									isBlank={blankVisibleColumns.has(key)}
									onRequestDelete={deleteColumn}
								/>
							);
						})}
					</SortableContext>
				</DndContext>
			</div>

			{/* Footer: Save button */}
			<div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 12 }}>
				<button
					onClick={onSave}
					disabled={!isDirty}
					aria-disabled={!isDirty}
					style={{
						padding: '8px 14px',
						borderRadius: 6,
						background: isDirty ? '#4f46e5' : '#9ca3af',
						color: '#fff',
						border: 'none',
						cursor: isDirty ? 'pointer' : 'not-allowed',
						opacity: isDirty ? 1 : 0.9,
						fontSize: 14
					}}
					title={isDirty ? 'Save changes' : 'No changes to save'}
				>
					Save
				</button>
			</div>
		</div>
	);
};
