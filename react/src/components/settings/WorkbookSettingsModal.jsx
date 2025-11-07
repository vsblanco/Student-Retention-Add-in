import React, { useEffect, useState } from "react";
import DeleteConfirmModal from './DeleteConfirmModal'; // <-- added

// fresh header button styles for modal
const modalBtnBase = {
	padding: '8px 10px',
	borderRadius: 10,
	border: 'none',
	cursor: 'pointer',
	boxShadow: '0 10px 30px rgba(79,70,229,0.08)',
	transition: 'transform 120ms ease, box-shadow 120ms ease',
	fontSize: 13,
	fontWeight: 600,
	marginRight: 8,
};
const modalBtnPrimary = { ...modalBtnBase, background: 'linear-gradient(90deg,#6366f1,#8b5cf6)', color: '#fff' };
const modalBtnGhost = { ...modalBtnBase, background: '#f8fafc', color: '#0f172a', boxShadow: 'none' };

// add a compact icon button variant for header icons
const modalIconBtn = { ...modalBtnGhost, padding: 6, width: 36, height: 36, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' };

// Accept docKey so the modal can explicitly read a named document-setting mapping
export default function WorkbookSettingsModal({ isOpen, onClose, docKey = 'workbookSettings' }) {
	const [loading, setLoading] = useState(false);
	const [data, setData] = useState(null);
	const [error, setError] = useState(null);
	// new: view toggle between tree explorer and raw JSON
	const [viewMode, setViewMode] = useState('tree'); // 'tree' | 'json'

	// NEW: delete confirm modal state
	const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
	const [deleting, setDeleting] = useState(false);

	useEffect(() => {
		if (isOpen) {
			loadSettings();
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [isOpen]);

	async function loadSettings() {
		setLoading(true);
		setError(null);
		try {
			const result = {
				workbookProperties: {},
				worksheets: [],
				namedItems: [],
				addinSettings: {},
				// this will hold the mapping stored under the provided docKey (if any)
				documentSettingsKey: docKey,
				documentSettings: null,
			};

			// Try to read Excel workbook properties, worksheets and named items via Excel.run if available
			if (window.Excel && Excel.run) {
				await Excel.run(async (context) => {
					const props = context.workbook.properties;
					const sheets = context.workbook.worksheets;
					const names = context.workbook.names;

					props.load(); // load all properties available on workbook.properties
					sheets.load("items/name");
					names.load("items/name,items/formula");

					await context.sync();

					// workbook properties -> plain object
					result.workbookProperties = {
						// map common properties if present
						title: props.title,
						subject: props.subject,
						author: props.author,
						company: props.company,
						manager: props.manager,
						category: props.category,
						keywords: props.keywords,
						comments: props.comments,
						// include raw props in case API has more
						_raw: props,
					};

					// worksheets
					result.worksheets = sheets.items.map((s) => s.name);

					// named items
					result.namedItems = names.items.map((n) => ({
						name: n.name,
						formula: n.formula,
					}));
				});
			} else {
				// Excel API not present (maybe running outside Excel). Provide a note.
				result._warning = "Excel JavaScript API not available in this environment.";
			}

			// Read add-in settings via Office.context.document.settings (Common API)
			if (window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				try {
					const settings = Office.context.document.settings;
					// settings is key/value store; Office.js doesn't provide an official enumerator,
					// but settings._data may exist in many hosts. We attempt a safe enumeration.
					const addinSettings = {};
					if (settings._data && typeof settings._data === "object") {
						Object.keys(settings._data).forEach((k) => {
							try {
								addinSettings[k] = settings.get
									? settings.get(k)
									: settings._data[k];
							} catch (e) {
								addinSettings[k] = "[error reading key]";
							}
						});
					} else {
						// Fallback: if you know specific keys, read them here
						// addinSettings.example = settings.get("example");
					}
					result.addinSettings = addinSettings;
					// Attempt to read the mapping stored under the provided docKey (if present)
					try {
						const mapping = typeof settings.get === 'function' ? settings.get(docKey) : (settings._data ? settings._data[docKey] : undefined);
						result.documentSettings = mapping ?? null;
					} catch (ex) {
						result.documentSettings = { _error: String(ex) };
					}
				} catch (e) {
					// ignore, include error message
					result.addinSettings = { _error: String(e) };
					// still attempt to read specific key if possible
					try {
						if (window.Office && Office.context && Office.context.document && Office.context.document.settings && typeof Office.context.document.settings.get === 'function') {
							result.documentSettings = Office.context.document.settings.get(docKey);
						}
					} catch (_) {
						/* ignore */
					}
				}
			} else {
				result.addinSettings = { _warning: "Office.context.document.settings not available." };
				result.documentSettings = { _warning: "Office.context.document.settings not available." };
			}

			setData(result);
		} catch (err) {
			setError(String(err));
			setData(null);
		} finally {
			setLoading(false);
		}
	}

	function close() {
		onClose && onClose();
	}

	async function copyJson() {
		if (!data) return;
		const text = JSON.stringify(data, null, 2);
		try {
			await navigator.clipboard.writeText(text);
			// minimal feedback could be added here
		} catch (e) {
			// fallback: create textarea and prompt user
			const el = document.createElement("textarea");
			el.value = text;
			document.body.appendChild(el);
			el.select();
			try {
				document.execCommand("copy");
			} catch (ex) {
				/* ignore */
			}
			document.body.removeChild(el);
		}
	}

	// NEW: internal delete logic (called after user confirms)
	async function deleteDocumentSettingsInternal() {
		setDeleting(true);
		try {
			// Prefer Office JS API remove + saveAsync if available
			if (window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				const settings = Office.context.document.settings;
				try {
					if (typeof settings.remove === 'function') {
						settings.remove(docKey);
					} else if (settings._data && Object.prototype.hasOwnProperty.call(settings._data, docKey)) {
						delete settings._data[docKey];
					}
					// attempt to persist change
					if (typeof settings.saveAsync === 'function') {
						await new Promise((resolve) => {
							try {
								settings.saveAsync(function () { resolve(null); });
							} catch (_) { resolve(null); }
						});
					} else if (typeof Office.context.document.settings.saveAsync === 'function') {
						await new Promise((resolve) => {
							try {
								Office.context.document.settings.saveAsync(function () { resolve(null); });
							} catch (_) { resolve(null); }
						});
					}
					// refresh view
					await loadSettings();
					return;
				} catch (e) {
					/* eslint-disable no-console */
					console.error("Error removing setting via Office API:", e);
					/* eslint-enable no-console */
				}
			}

			// fallback: clear local value and refresh
			setData((d) => d ? { ...d, documentSettings: null } : d);
		} finally {
			setDeleting(false);
			setDeleteConfirmOpen(false);
		}
	}

	if (!isOpen) return null;

	// --- New: small Tree/Explorer components ---
	function IconFolder({ open = false }) {
		return (
			<svg width="16" height="16" viewBox="0 0 24 24" fill="none" style={{ display: 'block' }} xmlns="http://www.w3.org/2000/svg">
				<path d="M3 7a2 2 0 0 1 2-2h4l2 2h6a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" fill={open ? '#f3f4f6' : 'none'} />
			</svg>
		);
	}
	function IconFile() {
		return (
			<svg width="14" height="14" viewBox="0 0 24 24" fill="none" style={{ display: 'block' }} xmlns="http://www.w3.org/2000/svg">
				<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
				<path d="M14 2v6h6" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
			</svg>
		);
	}
	function Caret({ open }) {
		return (
			<svg width="12" height="12" viewBox="0 0 24 24" style={{ transform: open ? 'rotate(90deg)' : 'none', transition: 'transform 120ms linear' }} fill="none" xmlns="http://www.w3.org/2000/svg">
				<path d="M8 5l8 7-8 7" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
			</svg>
		);
	}

	function TreeNode({ label, icon, defaultOpen = false, children, right }) {
		const [open, setOpen] = useState(defaultOpen);
		return (
			<div>
				<div
					onClick={() => setOpen(o => !o)}
					style={{
						display: 'flex',
						alignItems: 'center',
						gap: 8,
						padding: '6px 8px',
						cursor: children ? 'pointer' : 'default',
						userSelect: 'none',
						borderRadius: 6,
						transition: 'background 120ms',
					}}
				>
					{children ? <Caret open={open} /> : <div style={{ width: 12 }} />}
					<div style={{ width: 18, height: 18, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#6b7280' }}>
						{icon}
					</div>
					<div style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: 13 }}>
						{label}
					</div>
					{right && <div style={{ marginLeft: 8, color: '#6b7280', fontSize: 12 }}>{right}</div>}
				</div>
				{children && open && <div style={{ marginLeft: 20, borderLeft: '1px dashed rgba(0,0,0,0.04)', paddingLeft: 8 }}>{children}</div>}
			</div>
		);
	}
	// --- End Tree components ---

	// Recursive renderer: show objects/arrays as nested TreeNode folders; primitives show as leaf with right value.
	function renderValueAsTree(label, value, path = '') {
		const seen = renderValueAsTree.__seen || (renderValueAsTree.__seen = new WeakSet());

		// shortPreview: produce a compact preview for an item used in array labels.
		const shortPreview = (v, depth = 0) => {
			try {
				if (v === null || typeof v !== 'object') {
					const s = String(v ?? '');
					return s.length > 40 ? s.slice(0, 37) + '...' : s;
				}
				if (Array.isArray(v)) {
					if (v.length === 0) return '[empty array]';
					// preview first element
					return shortPreview(v[0], depth + 1);
				}
				// object: prefer a primitive property, otherwise first array element of any property
				const vals = Object.values(v);
				for (let i = 0; i < vals.length; i++) {
					if (vals[i] === null || typeof vals[i] !== 'object') {
						return shortPreview(vals[i], depth + 1);
					}
				}
				for (let i = 0; i < vals.length; i++) {
					if (Array.isArray(vals[i]) && vals[i].length > 0) {
						return shortPreview(vals[i][0], depth + 1);
					}
				}
				// fallback
				return '{object}';
			} catch (e) {
				return '{...}';
			}
		};

		const makeNode = (lbl, val, p) => {
			// primitives
			if (val === null || typeof val !== 'object') {
				return <TreeNode key={p} label={lbl} icon={<IconFile />} children={null} right={String(val ?? '')} />;
			}

			// circular guard
			if (seen.has(val)) {
				return <TreeNode key={p} label={lbl} icon={<IconFile />} children={null} right={'[circular]'} />;
			}
			seen.add(val);

			// array: show elements as "number: preview" (1-based index)
			if (Array.isArray(val)) {
				return (
					<TreeNode key={p} label={`${lbl} [array]`} icon={<IconFolder />}>
						{val.length === 0 ? (
							<div style={{ padding: 6, color: '#6b7280' }}>Empty array</div>
						) : (
							val.map((it, i) => {
								const previewLabel = `${i + 1}: ${shortPreview(it)}`;
								return makeNode(previewLabel, it, `${p}.${i}`);
							})
						)}
					</TreeNode>
				);
			}

			// object
			const entries = Object.entries(val);
			return (
				<TreeNode key={p} label={`${lbl} {object}`} icon={<IconFolder />}>
					{entries.length === 0 ? <div style={{ padding: 6, color: '#6b7280' }}>Empty object</div> :
						entries.map(([k, v]) => makeNode(k, v, `${p}.${k}`))}
				</TreeNode>
			);
		};

		try {
			return makeNode(label, value, path || label);
		} finally {
			// clear seen to allow subsequent independent renders
			renderValueAsTree.__seen = new WeakSet();
		}
	}

	return (
		<div style={{
			position: "fixed",
			inset: 0,
			background: "rgba(0,0,0,0.4)",
			display: "flex",
			alignItems: "center",
			justifyContent: "center",
			zIndex: 9999,
		}}>
			<div style={{
				width: "90%",
				maxWidth: 900,
				maxHeight: "90%",
				overflow: "auto",
				background: "#fff",
				borderRadius: 6,
				padding: 16,
				boxShadow: "0 6px 24px rgba(0,0,0,0.2)",
				fontFamily: "Segoe UI, system-ui, sans-serif",
				fontSize: 13,
			}}>
				<div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
					<h3 style={{ margin: 0 }}>Workbook Explorer</h3>
					<div>
						{/* toggle between tree view and json view (moved left) */}
						<button
							onClick={() => setViewMode(m => (m === 'tree' ? 'json' : 'tree'))}
							style={modalBtnGhost}
							title={viewMode === 'tree' ? 'Switch to JSON view' : 'Switch to tree view'}
							aria-label={viewMode === 'tree' ? 'Switch to JSON view' : 'Switch to tree view'}
						>
							{viewMode === 'tree' ? 'JSON View' : 'Tree View'}
						</button>
						{/* refresh icon button */}
						<button onClick={loadSettings} style={modalIconBtn} title="Refresh" aria-label="Refresh workbook settings">
							<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
								<path d="M21 12a9 9 0 1 1-2.36-6.36" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
								<path d="M21 3v6h-6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
							</svg>
						</button>

						{/* Copy JSON only visible in JSON view */}
						{viewMode === 'json' && <button onClick={copyJson} style={modalBtnPrimary}>Copy JSON</button>}
						{/* close as X icon */}
						<button onClick={close} style={modalIconBtn} title="Close" aria-label="Close workbook explorer">
							<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
								<path d="M18 6L6 18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
								<path d="M6 6l12 12" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
							</svg>
						</button>
					</div>
				</div>

				{loading && <div>Loading workbook settings...</div>}
				{error && <div style={{ color: "red" }}>Error: {error}</div>}

				{!loading && data && (
					viewMode === 'tree' ? (
						<div style={{ display: 'grid', gap: 8 }}>
							{/* Workbook Properties */}
							<TreeNode label="Workbook Properties" icon={<IconFolder />} defaultOpen={false}>
								{data.workbookProperties && Object.keys(data.workbookProperties).length > 0 ? (
									Object.entries(data.workbookProperties).map(([k, v]) => renderValueAsTree(k, v, `workbookProps.${k}`))
								) : (
									<div style={{ padding: 6, color: '#6b7280' }}>No properties available</div>
								)}
							</TreeNode>

							{/* Worksheets */}
							<TreeNode label={`Worksheets (${(data.worksheets && data.worksheets.length) || 0})`} icon={<IconFolder />}>
								{Array.isArray(data.worksheets) && data.worksheets.length > 0 ? (
									data.worksheets.map((name, idx) => (
										<TreeNode key={idx} label={name} icon={<IconFile />} />
									))
								) : (
									<div style={{ padding: 6, color: '#6b7280' }}>No worksheets found</div>
								)}
							</TreeNode>

							{/* Document-level mapping stored under provided docKey */}
							<TreeNode
								label={`${data.documentSettingsKey || docKey}`}
								icon={<IconFolder />}
								right={
									<button
										onClick={(e) => { e.stopPropagation(); setDeleteConfirmOpen(true); }}
										title="Delete mapping"
										style={{
											border: 'none',
											background: 'transparent',
											cursor: 'pointer',
											padding: 6,
											display: 'inline-flex',
											alignItems: 'center',
											justifyContent: 'center',
											color: '#ef4444',
											borderRadius: 6,
										}}
									>
										<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
											<path d="M3 6h18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
											<path d="M8 6v14a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2V6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
											<path d="M10 11v6M14 11v6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
											<path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
										</svg>
									</button>
								}
							>
								{data.documentSettings && typeof data.documentSettings === 'object' && Object.keys(data.documentSettings).length > 0 ? (
									Object.entries(data.documentSettings).map(([k, v]) => renderValueAsTree(k, v, `doc.${k}`))
								) : (
									<div style={{ padding: 6, color: '#6b7280' }}>{data.documentSettings === null ? 'No mapping found for this key' : String(data.documentSettings ?? '')}</div>
								)}
							</TreeNode>
						</div>
					) : (
						// JSON view
						<div>
							<pre style={{ margin: 6, padding: 8, background: '#f9fafb', borderRadius: 6, overflow: 'auto', maxHeight: 480, fontSize: 12 }}>
								{JSON.stringify(data, null, 2)}
							</pre>
						</div>
					)
				)}

				{!loading && !data && !error && (
					<div>No data available. Click Refresh.</div>
				)}
			</div>

			{/* Delete confirm modal */}
			<DeleteConfirmModal
				isOpen={deleteConfirmOpen}
				title="Delete document mapping"
				message={`Delete document setting "${docKey}"? This cannot be undone.`}
				confirmLabel={deleting ? 'Deleting...' : 'Delete'}
				onConfirm={deleteDocumentSettingsInternal}
				onCancel={() => setDeleteConfirmOpen(false)}
			/>
		</div>
	);
}