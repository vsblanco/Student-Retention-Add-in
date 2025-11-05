import React, { useState, useRef, useEffect } from 'react';
import { X, Plus } from 'lucide-react';

const SettingsModal = ({
	// array modal props
	modalOpen,
	modalSetting,
	modalArray = [],
	setModalArray = () => {},
	closeModal = () => {},
	saveModal = () => {},

	// selections modal props
	selectionsModalOpen,
	selectionsModalSetting,
	selectionsAvailable = [],
	selectionsChosen = [],
	selectionsFilter = '',
	setSelectionsFilter = () => {},
	closeSelectionsModal = () => {},
	saveSelectionsModal = () => {},
	moveToChosen = () => {},
	moveToAvailable = () => {},
}) => {
	// add drag state
	const [dragIndex, setDragIndex] = useState(null);
	const [dragOverIndex, setDragOverIndex] = useState(null);

	// highlight state for recently-moved row
	const [justMovedIndex, setJustMovedIndex] = useState(null);
	const highlightTimerRef = useRef(null);

	useEffect(() => {
		return () => {
			if (highlightTimerRef.current) {
				clearTimeout(highlightTimerRef.current);
			}
		};
	}, []);

	// list scroll flags for array modal
	const listScrollable = Array.isArray(modalArray) && modalArray.length > 4;

	// drag handlers for reordering
	const onDragStart = (e, idx) => {
		setDragIndex(idx);
		// some browsers require data set for drag to start
		try { e.dataTransfer.setData('text/plain', String(idx)); } catch (err) {}
		e.dataTransfer.effectAllowed = 'move';
	};
	const onDragOver = (e, idx) => {
		e.preventDefault(); // allow drop
		if (dragOverIndex !== idx) setDragOverIndex(idx);
	};
	const onDrop = (e, idx) => {
		e.preventDefault();
		if (dragIndex === null) return;
		if (dragIndex === idx) {
			setDragIndex(null);
			setDragOverIndex(null);
			return;
		}
		setModalArray(prev => {
			const copy = [...prev];
			const [moved] = copy.splice(dragIndex, 1);
			// insert at target index (works for both directions)
			copy.splice(idx, 0, moved);
			return copy;
		});
		// show quick highlight where the row landed
		setJustMovedIndex(idx);
		if (highlightTimerRef.current) clearTimeout(highlightTimerRef.current);
		highlightTimerRef.current = setTimeout(() => setJustMovedIndex(null), 800);

		setDragIndex(null);
		setDragOverIndex(null);
	};
	const onDragEnd = () => {
		setDragIndex(null);
		setDragOverIndex(null);
	};

	// selections derived values
	// If the setting provides a key (e.g. 'hidden'), derive chosen/available from selectionsAvailable using that key.
	const keyName = selectionsModalSetting?.key;
	let derivedAvailable = Array.isArray(selectionsAvailable) ? [...selectionsAvailable] : [];
	let derivedChosen = Array.isArray(selectionsChosen) ? [...selectionsChosen] : [];

	if (keyName && Array.isArray(selectionsAvailable)) {
		derivedChosen = selectionsAvailable.filter(item => !!item?.[keyName]);
		derivedAvailable = selectionsAvailable.filter(item => !item?.[keyName]);
	}

	const filteredAvailable = derivedAvailable.filter(item =>
		String(item.label ?? '').toLowerCase().includes(String(selectionsFilter ?? '').toLowerCase())
	);
	const availableScrollable = filteredAvailable.length > 4;
	const chosenScrollable = derivedChosen.length > 4;

	return (
		<>
			{/* Array configure modal OR editableArray modal */}
			{modalOpen && modalSetting?.type === 'editableArray' ? (
				<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
					<div style={{ width: 'min(820px, 96%)', maxHeight: '86vh', overflow: 'auto', background: '#fff', borderRadius: 8, padding: 16, boxSizing: 'border-box' }}>
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
						/>
					</div>
				</div>
			) : modalOpen && (
				<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
					<div style={{ width: 'min(720px, 96%)', maxHeight: '80vh', overflow: 'auto', background: '#fff', borderRadius: 8, padding: 16, boxSizing: 'border-box' }}>
						<div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
							<h3 style={{ margin: 0 }}>{modalSetting?.label || 'Configure array'}</h3>
							<button
								onClick={closeModal}
								style={{ background: 'transparent', border: 'none', cursor: 'pointer', padding: 6, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
								aria-label="Close"
							>
								<X size={16} />
							</button>
						</div>

						<div style={{ display: 'grid', gap: 8 }}>
							{/* list becomes scrollable when more than 4 entries */}
							<div style={{
								display: 'grid',
								gap: 8,
								maxHeight: listScrollable ? '240px' : 'auto',
								overflowY: listScrollable ? 'auto' : 'visible',
								overflowX: 'hidden',
								paddingRight: listScrollable ? 8 : 0,
								boxSizing: 'border-box'
							}}>
								{modalArray.map((item, idx) => (
									// make row a drop target only; handle is draggable
									<div
										key={idx}
										onDragOver={e => onDragOver(e, idx)}
										onDrop={e => onDrop(e, idx)}
										style={{
											display: 'flex',
											gap: 8,
											alignItems: 'center',
											border: dragIndex === idx ? '2px dashed #4f46e5' : (dragOverIndex === idx ? '2px dashed #93c5fd' : 'none'),
											padding: dragIndex === idx || dragOverIndex === idx ? 6 : 0,
											borderRadius: 6,
											// highlight animation: temporary warm background + subtle transition
											background: dragIndex === idx ? '#f8fafc' : (justMovedIndex === idx ? '#bbdefc' : 'transparent'),
											transition: 'background-color 360ms ease, box-shadow 360ms ease'
										}}
									>
										{/* draggable handle (click-and-hold here to drag) */}
										<button
											type="button"
											draggable
											onDragStart={e => onDragStart(e, idx)}
											onDragEnd={onDragEnd}
											title="Drag to reorder"
											aria-label={`Drag item ${idx + 1}`}
											style={{
												display: 'inline-flex',
												alignItems: 'center',
												justifyContent: 'center',
												width: 28,
												height: 28,
												borderRadius: 6,
												background: 'transparent',
												border: '1px solid transparent',
												cursor: dragIndex === idx ? 'grabbing' : 'grab',
												padding: 4,
												flexShrink: 0
											}}
										>
											{/* grip icon: six dots */}
											<svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden>
												<circle cx="3" cy="3" r="1" fill="#6b7280"/>
												<circle cx="7" cy="3" r="1" fill="#6b7280"/>
												<circle cx="11" cy="3" r="1" fill="#6b7280"/>
												<circle cx="3" cy="7" r="1" fill="#6b7280"/>
												<circle cx="7" cy="7" r="1" fill="#6b7280"/>
												<circle cx="11" cy="7" r="1" fill="#6b7280"/>
											</svg>
										</button>

										<input
											type="text"
											value={item.edit === 'name' ? (item.name ?? '') : (Array.isArray(item.alias) ? item.alias.join(', ') : (item.alias ?? ''))}
											onChange={e => {
												const copy = [...modalArray];
												if (copy[idx].edit === 'name') {
													copy[idx] = { ...copy[idx], name: e.target.value };
												} else {
													const parsed = String(e.target.value)
														.split(',')
														.map(s => s.trim())
														.filter(Boolean);
													copy[idx] = { ...copy[idx], alias: parsed };
												}
												setModalArray(copy);
											}}
											placeholder={item.edit === 'name' ? 'name' : 'aliases (comma separated)'}
											style={{ flex: 1, minWidth: 0, padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb' }}
											title={item.static ? 'Static flag is set (this is metadata); value remains editable' : undefined}
										/>

										{/* aliases toggle */}
										{(() => {
											const iconColor = item.edit === 'alias' ? '#2563eb' : '#60a5fa';
											const bg = item.edit === 'alias' ? '#e8f0ff' : '#f3f8ff';
											return (
												<button
													onClick={() => {
														const copy = [...modalArray];
														copy[idx] = { ...copy[idx], edit: copy[idx].edit === 'name' ? 'alias' : 'name' };
														setModalArray(copy);
													}}
													title={item.edit === 'name' ? 'Edit aliases' : 'Edit name'}
													aria-label={item.edit === 'name' ? 'Edit aliases' : 'Edit name'}
													style={{
														display: 'inline-flex',
														alignItems: 'center',
														justifyContent: 'center',
														width: 34,
														height: 34,
														borderRadius: 6,
														background: bg,
														border: '1px solid rgba(37,99,235,0.12)',
														cursor: 'pointer',
														color: iconColor,
														transition: 'color 120ms ease, background-color 120ms ease, box-shadow 120ms ease'
													}}
												>
													<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block' }}>
														<path d="M20.59 13.41L11 3.83A2 2 0 0 0 9.59 3H5a2 2 0 0 0-2 2v4.59c0 .53.21 1.04.59 1.41l9.59 9.59a2 2 0 0 0 2.83 0l4.59-4.59a2 2 0 0 0 0-2.83z" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
														<circle cx="7.5" cy="7.5" r="1.5" fill="currentColor"/>
													</svg>
												</button>
											);
										})()}

										{/* static toggle */}
										<button
											onClick={() => {
												const copy = [...modalArray];
												copy[idx] = { ...copy[idx], static: !copy[idx].static };
												setModalArray(copy);
											}}
											title={item.static ? 'Unset static (make editable)' : 'Set static (lock)'}
											aria-label={item.static ? 'Unset static' : 'Set static'}
											aria-pressed={!!item.static}
											style={{
												display: 'inline-flex',
												alignItems: 'center',
												justifyContent: 'center',
												width: 34,
												height: 34,
												borderRadius: 6,
												background: item.static ? '#fff7ed' : '#f3f4f6',
												border: '1px solid rgba(0,0,0,0.06)',
												cursor: 'pointer',
												color: item.static ? '#b45309' : '#6b7280'
											}}
										>
											{item.static ? (
												<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
													<rect x="3" y="11" width="18" height="10" rx="2" stroke="currentColor" strokeWidth="1.2"/>
													<path d="M7 11V8a5 5 0 0 1 10 0v3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
												</svg>
											) : (
												<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
													<rect x="3" y="11" width="18" height="10" rx="2" stroke="currentColor" strokeWidth="1.2"/>
													<path d="M16 11V8a4 4 0 0 0-8 0" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
												</svg>
											)}
										</button>

										<button
											onClick={() => setModalArray(prev => prev.filter((_, i) => i !== idx))}
											style={{
												padding: '6px 8px',
												borderRadius: 6,
												background: '#fff5f5',
												border: '1px solid rgba(239,68,68,0.12)',
												color: '#ef4444',
												cursor: 'pointer',
												transition: 'background-color 120ms ease, color 120ms ease, transform 120ms ease'
											}}
											aria-label={`Remove item ${idx + 1}`}
										>
											<X size={14} color="#ef4444" />
										</button>
									</div>
								))}
							</div>

							<button
								onClick={() => setModalArray(prev => [...prev, { name: '', alias: [], edit: 'name', static: false }])}
								style={{ padding: '8px 10px', borderRadius: 6, width: 'max-content' }}
							>
								Add item
							</button>
						</div>

						<div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 12 }}>
							<button onClick={closeModal} style={{ padding: '8px 10px', borderRadius: 6 }}>Cancel</button>
							<button onClick={saveModal} style={{ padding: '8px 10px', borderRadius: 6, background: '#4f46e5', color: '#fff', border: 'none' }}>Save</button>
						</div>
					</div>
				</div>
			)}

			{/* Selections modal */}
			{selectionsModalOpen && (
				<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
					<div style={{ width: 'min(900px, 96%)', maxHeight: '86vh', overflow: 'auto', background: '#fff', borderRadius: 8, padding: 16, boxSizing: 'border-box' }}>
						<div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
							<h3 style={{ margin: 0 }}>{selectionsModalSetting?.label || 'Configure'}</h3>
							<button
								onClick={closeSelectionsModal}
								style={{ background: 'transparent', border: 'none', cursor: 'pointer', padding: 6, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
								aria-label="Close"
							>
								<X size={16} />
							</button>
						</div>

						<div style={{ display: 'grid', gridTemplateColumns: '1fr 12px 1fr', gap: 12, alignItems: 'start' }}>
							{/* Available bank */}
							<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 12, background: '#fafafa' }}>
								<div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
									<input
										type="search"
										placeholder="Search available..."
										value={selectionsFilter}
										onChange={e => setSelectionsFilter(e.target.value)}
										style={{ flex: 1, padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb' }}
										aria-label="Filter available selections"
									/>
								</div>

								<div style={{
									display: 'grid',
									gap: 8,
									maxHeight: availableScrollable ? '240px' : 'auto',
									overflowY: availableScrollable ? 'auto' : 'visible',
									overflowX: 'hidden',
									paddingRight: availableScrollable ? 8 : 0,
									boxSizing: 'border-box'
								}}>
									{filteredAvailable.length === 0 && (
										<div style={{ color: '#6b7280', fontSize: 13 }}>No available items</div>
									)}
									{filteredAvailable.map(item => (
										<div key={item.key} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
											<div style={{ overflow: 'hidden', minWidth: 0, textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.label}</div>
											<button
												onClick={() => moveToChosen(item)}
												aria-label={`Add ${item.label}`}
												title={`Add ${item.label}`}
												style={{ padding: 6, borderRadius: 6, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
											>
												<Plus size={16} />
											</button>
										</div>
									))}
								</div>
							</div>

							{/* spacer */}
							<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }} aria-hidden>
								<div style={{ writingMode: 'vertical-rl', transform: 'rotate(180deg)', color: '#9ca3af' }}>Selected →</div>
							</div>

							{/* Chosen bank */}
							<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 12, background: '#fff' }}>
								<div style={{ marginBottom: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
									<div style={{ fontWeight: 600 }}>Chosen</div>
									<div style={{ fontSize: 13, color: '#6b7280' }}>{derivedChosen.length} selected</div>
								</div>

								<div style={{
									display: 'grid',
									gap: 8,
									maxHeight: chosenScrollable ? '240px' : 'auto',
									overflowY: chosenScrollable ? 'auto' : 'visible',
									overflowX: 'hidden',
									paddingRight: chosenScrollable ? 8 : 0,
									boxSizing: 'border-box'
								}}>
									{derivedChosen.length === 0 && (
										<div style={{ color: '#6b7280', fontSize: 13 }}>No items chosen</div>
									)}
									{derivedChosen.map(item => (
										<div key={item.key} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
											<div style={{ overflow: 'hidden', minWidth: 0, textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.label}</div>
											<button
												onClick={() => moveToAvailable(item)}
												aria-label={`Remove ${item.label}`}
												title={`Remove ${item.label}`}
												style={{
													padding: '6px 8px',
													borderRadius: 6,
													background: '#fff5f5',
													border: '1px solid rgba(239,68,68,0.12)',
													color: '#ef4444',
													cursor: 'pointer',
													display: 'inline-flex',
													alignItems: 'center',
													justifyContent: 'center'
												}}
											>
												<X size={14} color="#ef4444" />
											</button>
										</div>
									))}
								</div>
							</div>
						</div>

						<div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 12 }}>
							<button onClick={closeSelectionsModal} style={{ padding: '8px 10px', borderRadius: 6 }}>Cancel</button>
							<button onClick={saveSelectionsModal} style={{ padding: '8px 10px', borderRadius: 6, background: '#4f46e5', color: '#fff', border: 'none' }}>Save</button>
						</div>
					</div>
				</div>
			)}
		</>
	);
};

export default SettingsModal;

// Add small helper component in the same file (keeps changes minimal)
const EditableArrayInner = ({ modalSetting, modalArray = [], setModalArray, closeModal, saveModal }) => {
	const [selectedIdx, setSelectedIdx] = useState(0);
	// editableMap: { [columnName]: { [optionName]: selectedTypeOrEmptyString } }
	const [editableMap, setEditableMap] = useState({});

	useEffect(() => {
		// initialize map from modalArray if it follows expected shape, otherwise default
		const map = {};
		// modalArray expected to be array of objects like: { column: 'Name', options: { 'Format': 'MM/DD/YYYY' } }
		if (Array.isArray(modalArray) && modalArray.length) {
			modalArray.forEach(entry => {
				if (entry && entry.column) {
					map[entry.column] = { ...(entry.options || {}) };
				}
			});
		}
		// ensure every choice exists in map
		(modalSetting?.choices || []).forEach(choice => {
			const key = String(choice.name ?? choice.label ?? choice).trim();
			if (!map[key]) map[key] = {};
		});
		setEditableMap(map);
		// pick previously-selected index if possible
		setSelectedIdx(prev => (modalSetting?.choices && modalSetting.choices[prev] ? prev : 0));
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [modalOpenOrSettingMarker() /* dummy dependency to silence lint; see helper below */]);

	// helper for effect dependency without introducing new props
	function modalOpenOrSettingMarker() {
		// combine relevant values to create a shallow marker — avoids adding modalArray to deps directly
		return JSON.stringify({
			choices: modalSetting?.choices?.length ?? 0,
			options: modalSetting?.options?.length ?? 0,
			modalArrayLength: (modalArray || []).length
		});
	}

	if (!modalSetting) return null;

	const choices = modalSetting.choices || [];
	const options = modalSetting.options || [];

	const currentChoice = choices[selectedIdx];
	const currentKey = String(currentChoice?.name ?? currentChoice?.label ?? currentChoice ?? '').trim();

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
		// transform editableMap into an array for setModalArray: [{ column, options: }]
		const out = Object.keys(editableMap).map(col => ({ column: col, options: { ...(editableMap[col] || {}) } }));
		setModalArray(out);
		saveModal && saveModal();
		closeModal && closeModal();
	};

	return (
		<div style={{ display: 'grid', gridTemplateColumns: '220px 12px 1fr', gap: 12 }}>
			{/* Choice selector (list) */}
			<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 8, background: '#fafafa', maxHeight: '56vh', overflowY: 'auto' }}>
				<div style={{ marginBottom: 8, fontWeight: 600 }}>Columns</div>
				{choices.length === 0 && <div style={{ color: '#6b7280', fontSize: 13 }}>No choices available</div>}
				{choices.map((c, i) => {
					const label = String(c.name ?? c.label ?? c).trim();
					const selected = i === selectedIdx;
					return (
						<button
							key={label + i}
							onClick={() => setSelectedIdx(i)}
							style={{
								display: 'block',
								width: '100%',
								textAlign: 'left',
								padding: '8px',
								marginBottom: 6,
								borderRadius: 6,
								background: selected ? '#eef2ff' : 'transparent',
								border: selected ? '1px solid rgba(79,70,229,0.12)' : '1px solid transparent',
								cursor: 'pointer'
							}}
						>
							<div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</div>
						</button>
					);
				})}
			</div>

			{/* spacer */}
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }} aria-hidden>
				<div style={{ writingMode: 'vertical-rl', transform: 'rotate(180deg)', color: '#9ca3af' }}>Options →</div>
			</div>

			{/* Options editor */}
			<div style={{ border: '1px solid #e6e7eb', borderRadius: 6, padding: 12, background: '#fff', maxHeight: '56vh', overflowY: 'auto' }}>
				<div style={{ marginBottom: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
					<div style={{ fontWeight: 600 }}>Options for: <span style={{ fontWeight: 700 }}>{currentKey}</span></div>
					<div style={{ fontSize: 13, color: '#6b7280' }}>{Object.keys(editableMap[currentKey] || {}).length} set</div>
				</div>

				<div style={{ display: 'grid', gap: 8 }}>
					{options.length === 0 && <div style={{ color: '#6b7280', fontSize: 13 }}>No options available</div>}
					{options.map(opt => {
						const optName = opt.option || opt.name || String(opt);
						const types = Array.isArray(opt.type) ? opt.type : [];
						const value = (editableMap[currentKey] && editableMap[currentKey][optName]) || '';
						return (
							<div key={optName} style={{ display: 'flex', gap: 8, alignItems: 'center', justifyContent: 'space-between' }}>
								<div style={{ minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{optName}</div>
								<select
									value={value}
									onChange={e => updateOption(optName, e.target.value)}
									style={{ padding: '6px 8px', borderRadius: 6, border: '1px solid #e6e7eb', minWidth: 160 }}
									aria-label={`Select ${optName} for ${currentKey}`}
								>
									<option value=''>None</option>
									{types.map(t => (
										<option key={t} value={t}>{t}</option>
									))}
								</select>
							</div>
						);
					})}
				</div>

				{/* Actions */}
				<div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 12 }}>
					<button onClick={closeModal} style={{ padding: '8px 10px', borderRadius: 6 }}>Cancel</button>
					<button onClick={onSave} style={{ padding: '8px 10px', borderRadius: 6, background: '#4f46e5', color: '#fff', border: 'none' }}>Save</button>
				</div>
			</div>
		</div>
	);
};
