import React, { useEffect, useRef, useState } from "react";
import DeleteConfirmModal from './DeleteConfirmModal';
import { TreeNode, IconFolder, IconFile, renderValueAsTree } from './workbookDebug/Tree';
import { isAuthorMatch } from './workbookDebug/allowlist';

const BRAND = '#145F82';

const modalBtnBase = {
	height: 32,
	padding: '0 12px',
	borderRadius: 8,
	border: 'none',
	cursor: 'pointer',
	transition: 'transform 120ms ease, box-shadow 120ms ease, background 120ms ease',
	fontSize: 13,
	fontWeight: 600,
	display: 'inline-flex',
	alignItems: 'center',
	justifyContent: 'center',
	gap: 6,
};
const modalBtnPrimary = { ...modalBtnBase, background: BRAND, color: '#fff' };
const modalBtnGhost = { ...modalBtnBase, background: '#f8fafc', color: '#0f172a' };
const modalBtnDanger = { ...modalBtnBase, background: '#fee2e2', color: '#991b1b' };
const modalIconBtn = { ...modalBtnGhost, padding: 0, width: 32 };

// Accept docKey so the modal can explicitly read a named document-setting mapping
export default function WorkbookSettingsModal({ isOpen, onClose, docKey = 'workbookSettings' }) {
	const [loading, setLoading] = useState(false);
	const [data, setData] = useState(null);
	const [error, setError] = useState(null);
	const [viewMode, setViewMode] = useState('tree'); // 'tree' | 'json'
	const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
	const [deleting, setDeleting] = useState(false);
	const [copied, setCopied] = useState(false);
	const [currentUserName, setCurrentUserName] = useState('');
	const copiedTimerRef = useRef(null);
	const closeBtnRef = useRef(null);

	useEffect(() => {
		try {
			const raw = localStorage.getItem('SSO_USER_INFO');
			if (raw) setCurrentUserName(JSON.parse(raw).name || '');
		} catch { /* ignore */ }
	}, []);

	useEffect(() => {
		if (isOpen) {
			loadSettings();
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [isOpen]);

	// Esc-to-close + initial focus management
	useEffect(() => {
		if (!isOpen) return;
		const handleKey = (e) => {
			if (e.key === 'Escape' && !deleteConfirmOpen) {
				e.stopPropagation();
				onClose && onClose();
			}
		};
		document.addEventListener('keydown', handleKey);
		const focusTimer = setTimeout(() => closeBtnRef.current && closeBtnRef.current.focus(), 0);
		return () => {
			document.removeEventListener('keydown', handleKey);
			clearTimeout(focusTimer);
		};
	}, [isOpen, deleteConfirmOpen, onClose]);

	useEffect(() => () => {
		if (copiedTimerRef.current) clearTimeout(copiedTimerRef.current);
	}, []);

	async function loadSettings() {
		setLoading(true);
		setError(null);
		try {
			const result = {
				workbookProperties: {},
				worksheets: [],
				namedItems: [],
				addinSettings: {},
				documentSettingsKey: docKey,
				documentSettings: null,
			};

			if (window.Excel && Excel.run) {
				await Excel.run(async (context) => {
					const props = context.workbook.properties;
					const sheets = context.workbook.worksheets;
					const names = context.workbook.names;

					props.load();
					sheets.load("items/name");
					names.load("items/name,items/formula");

					await context.sync();

					result.workbookProperties = {
						title: props.title,
						subject: props.subject,
						author: props.author,
						company: props.company,
						manager: props.manager,
						category: props.category,
						keywords: props.keywords,
						comments: props.comments,
					};
					result.worksheets = sheets.items.map((s) => s.name);
					result.namedItems = names.items.map((n) => ({ name: n.name, formula: n.formula }));
				});
			} else {
				result._warning = "Excel JavaScript API not available in this environment.";
			}

			if (window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				try {
					const settings = Office.context.document.settings;
					const addinSettings = {};
					if (settings._data && typeof settings._data === "object") {
						Object.keys(settings._data).forEach((k) => {
							try {
								addinSettings[k] = settings.get ? settings.get(k) : settings._data[k];
							} catch {
								addinSettings[k] = "[error reading key]";
							}
						});
					}
					result.addinSettings = addinSettings;
					try {
						const mapping = typeof settings.get === 'function'
							? settings.get(docKey)
							: (settings._data ? settings._data[docKey] : undefined);
						result.documentSettings = mapping ?? null;
					} catch (ex) {
						result.documentSettings = { _error: String(ex) };
					}
				} catch (e) {
					result.addinSettings = { _error: String(e) };
					try {
						if (window.Office && Office.context && Office.context.document && Office.context.document.settings && typeof Office.context.document.settings.get === 'function') {
							result.documentSettings = Office.context.document.settings.get(docKey);
						}
					} catch {
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
		} catch {
			const el = document.createElement("textarea");
			el.value = text;
			document.body.appendChild(el);
			el.select();
			try { document.execCommand("copy"); } catch { /* ignore */ }
			document.body.removeChild(el);
		}
		setCopied(true);
		if (copiedTimerRef.current) clearTimeout(copiedTimerRef.current);
		copiedTimerRef.current = setTimeout(() => setCopied(false), 1500);
	}

	async function deleteDocumentSettingsInternal() {
		setDeleting(true);
		try {
			if (window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				const settings = Office.context.document.settings;
				try {
					if (typeof settings.remove === 'function') {
						settings.remove(docKey);
					} else if (settings._data && Object.prototype.hasOwnProperty.call(settings._data, docKey)) {
						delete settings._data[docKey];
					}
					if (typeof settings.saveAsync === 'function') {
						await new Promise((resolve) => {
							try { settings.saveAsync(() => resolve(null)); } catch { resolve(null); }
						});
					}
					await loadSettings();
					return;
				} catch (e) {
					/* eslint-disable no-console */
					console.error("Error removing setting via Office API:", e);
					/* eslint-enable no-console */
				}
			}
			setData((d) => d ? { ...d, documentSettings: null } : d);
		} finally {
			setDeleting(false);
			setDeleteConfirmOpen(false);
		}
	}

	if (!isOpen) return null;

	const hasDocumentMapping = data && data.documentSettings && typeof data.documentSettings === 'object' && Object.keys(data.documentSettings).length > 0;
	const canReset = isAuthorMatch(data?.workbookProperties?.author, currentUserName);
	const showFooter = (viewMode === 'json') || (hasDocumentMapping && canReset);

	// Segmented view toggle
	const SegmentedToggle = () => (
		<div style={{
			position: 'relative',
			display: 'flex',
			background: '#e2e8f0',
			borderRadius: 999,
			padding: 2,
			height: 32,
			width: 160,
		}}>
			<button
				type="button"
				onClick={() => setViewMode('tree')}
				style={{
					position: 'relative', zIndex: 1, flex: 1, border: 'none', background: 'transparent',
					borderRadius: 999, cursor: 'pointer', fontSize: 12, fontWeight: 600,
					color: viewMode === 'tree' ? '#fff' : '#64748b', transition: 'color 120ms ease',
				}}
			>
				Tree
			</button>
			<button
				type="button"
				onClick={() => setViewMode('json')}
				style={{
					position: 'relative', zIndex: 1, flex: 1, border: 'none', background: 'transparent',
					borderRadius: 999, cursor: 'pointer', fontSize: 12, fontWeight: 600,
					color: viewMode === 'json' ? '#fff' : '#64748b', transition: 'color 120ms ease',
				}}
			>
				JSON
			</button>
			<span style={{
				position: 'absolute', top: 2, bottom: 2,
				left: viewMode === 'tree' ? 2 : 'calc(50%)',
				width: 'calc(50% - 2px)',
				borderRadius: 999, background: BRAND,
				transition: 'left 200ms ease',
			}} />
		</div>
	);

	return (
		<div
			onClick={(e) => { if (e.target === e.currentTarget) close(); }}
			style={{
				position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)",
				display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999,
			}}
		>
			<div style={{
				width: "90%", maxWidth: 900, maxHeight: "90%",
				background: "#fff", borderRadius: 10,
				boxShadow: "0 6px 24px rgba(0,0,0,0.2)",
				fontFamily: "Segoe UI, system-ui, sans-serif", fontSize: 13,
				display: 'flex', flexDirection: 'column',
			}}>
				{/* Sticky header */}
				<div style={{
					display: 'flex', justifyContent: 'space-between', alignItems: 'center',
					gap: 12, padding: '12px 16px',
					borderBottom: '1px solid #e5e7eb', flex: '0 0 auto',
				}}>
					<h3 style={{ margin: 0, fontSize: 16, lineHeight: '32px' }}>Workbook Debug</h3>
					<div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
						<SegmentedToggle />
						<button onClick={loadSettings} style={modalIconBtn} title="Refresh" aria-label="Refresh workbook debug view">
							<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
								<path d="M21 12a9 9 0 1 1-2.36-6.36" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
								<path d="M21 3v6h-6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
							</svg>
						</button>
						<button ref={closeBtnRef} onClick={close} style={modalIconBtn} title="Close" aria-label="Close workbook debug view">
							<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
								<path d="M18 6L6 18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
								<path d="M6 6l12 12" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
							</svg>
						</button>
					</div>
				</div>

				{/* Scrollable content */}
				<div style={{ flex: '1 1 auto', overflow: 'auto', padding: 16 }}>
					{loading && <div>Loading workbook debug data...</div>}
					{error && <div style={{ color: "red" }}>Error: {error}</div>}

					{!loading && data && (
						viewMode === 'tree' ? (
							<div style={{ display: 'grid', gap: 8 }}>
								<TreeNode label="Workbook Properties" icon={<IconFolder />} defaultOpen={false}>
									{data.workbookProperties && Object.keys(data.workbookProperties).length > 0
										? Object.entries(data.workbookProperties).map(([k, v]) => renderValueAsTree(k, v, `workbookProps.${k}`))
										: <div style={{ padding: 6, color: '#6b7280' }}>No properties available</div>}
								</TreeNode>

								<TreeNode label={`Worksheets (${(data.worksheets && data.worksheets.length) || 0})`} icon={<IconFolder />}>
									{Array.isArray(data.worksheets) && data.worksheets.length > 0
										? data.worksheets.map((name, idx) => <TreeNode key={idx} label={name} icon={<IconFile />} />)
										: <div style={{ padding: 6, color: '#6b7280' }}>No worksheets found</div>}
								</TreeNode>

								<TreeNode label={`${data.documentSettingsKey || docKey}`} icon={<IconFolder />}>
									{hasDocumentMapping
										? Object.entries(data.documentSettings).map(([k, v]) => renderValueAsTree(k, v, `doc.${k}`))
										: <div style={{ padding: 6, color: '#6b7280' }}>{data.documentSettings === null ? 'No mapping found for this key' : String(data.documentSettings ?? '')}</div>}
								</TreeNode>
							</div>
						) : (
							<pre style={{ margin: 0, padding: 8, background: '#f9fafb', borderRadius: 6, overflow: 'auto', fontSize: 12 }}>
								{JSON.stringify(data, null, 2)}
							</pre>
						)
					)}

					{!loading && !data && !error && <div>No data available. Click Refresh.</div>}
				</div>

				{/* Sticky footer (conditional) */}
				{showFooter && (
					<div style={{
						display: 'flex', justifyContent: 'space-between', alignItems: 'center',
						gap: 8, padding: '12px 16px',
						borderTop: '1px solid #e5e7eb', background: '#fafafa',
						borderRadius: '0 0 10px 10px', flex: '0 0 auto',
					}}>
						<div>
							{viewMode === 'tree' && hasDocumentMapping && canReset && (
								<button
									onClick={() => setDeleteConfirmOpen(true)}
									style={modalBtnDanger}
									title={`Delete the "${docKey}" document mapping`}
								>
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
										<path d="M3 6h18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
										<path d="M8 6v14a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2V6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
										<path d="M10 11v6M14 11v6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
										<path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
									</svg>
									Reset workbook defaults
								</button>
							)}
						</div>
						<div>
							{viewMode === 'json' && (
								<button onClick={copyJson} style={copied ? { ...modalBtnPrimary, background: '#16a34a' } : modalBtnPrimary}>
									{copied ? 'Copied!' : 'Copy JSON'}
								</button>
							)}
						</div>
					</div>
				)}
			</div>

			<DeleteConfirmModal
				isOpen={deleteConfirmOpen}
				title="Reset workbook defaults"
				message={`Delete the "${docKey}" document mapping? This cannot be undone.`}
				confirmLabel={deleting ? 'Deleting...' : 'Reset'}
				onConfirm={deleteDocumentSettingsInternal}
				onCancel={() => setDeleteConfirmOpen(false)}
			/>
		</div>
	);
}
