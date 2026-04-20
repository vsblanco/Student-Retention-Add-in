import React, { useState } from 'react';

export function IconFolder({ open = false }) {
	return (
		<svg width="16" height="16" viewBox="0 0 24 24" fill="none" style={{ display: 'block' }} xmlns="http://www.w3.org/2000/svg">
			<path d="M3 7a2 2 0 0 1 2-2h4l2 2h6a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" fill={open ? '#f3f4f6' : 'none'} />
		</svg>
	);
}

export function IconFile() {
	return (
		<svg width="14" height="14" viewBox="0 0 24 24" fill="none" style={{ display: 'block' }} xmlns="http://www.w3.org/2000/svg">
			<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
			<path d="M14 2v6h6" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
		</svg>
	);
}

export function Caret({ open }) {
	return (
		<svg width="12" height="12" viewBox="0 0 24 24" style={{ transform: open ? 'rotate(90deg)' : 'none', transition: 'transform 120ms linear' }} fill="none" xmlns="http://www.w3.org/2000/svg">
			<path d="M8 5l8 7-8 7" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
		</svg>
	);
}

export function TreeNode({ label, icon, defaultOpen = false, children, right }) {
	const [open, setOpen] = useState(defaultOpen);
	const [hover, setHover] = useState(false);
	return (
		<div>
			<div
				onClick={() => setOpen(o => !o)}
				onMouseEnter={() => setHover(true)}
				onMouseLeave={() => setHover(false)}
				style={{
					display: 'flex',
					alignItems: 'center',
					gap: 8,
					padding: '6px 8px',
					cursor: children ? 'pointer' : 'default',
					userSelect: 'none',
					borderRadius: 6,
					background: hover ? 'rgba(15,23,42,0.04)' : 'transparent',
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

// Recursive renderer: show objects/arrays as nested TreeNode folders; primitives show as leaf with right value.
export function renderValueAsTree(label, value, path = '') {
	const seen = new WeakSet();

	const shortPreview = (v) => {
		try {
			if (v === null || typeof v !== 'object') {
				const s = String(v ?? '');
				return s.length > 40 ? s.slice(0, 37) + '...' : s;
			}
			if (Array.isArray(v)) {
				if (v.length === 0) return '[empty array]';
				return shortPreview(v[0]);
			}
			const vals = Object.values(v);
			for (let i = 0; i < vals.length; i++) {
				if (vals[i] === null || typeof vals[i] !== 'object') {
					return shortPreview(vals[i]);
				}
			}
			for (let i = 0; i < vals.length; i++) {
				if (Array.isArray(vals[i]) && vals[i].length > 0) {
					return shortPreview(vals[i][0]);
				}
			}
			return '{object}';
		} catch {
			return '{...}';
		}
	};

	const makeNode = (lbl, val, p) => {
		if (val === null || typeof val !== 'object') {
			return <TreeNode key={p} label={lbl} icon={<IconFile />} right={String(val ?? '')} />;
		}
		if (seen.has(val)) {
			return <TreeNode key={p} label={lbl} icon={<IconFile />} right={'[circular]'} />;
		}
		seen.add(val);

		if (Array.isArray(val)) {
			return (
				<TreeNode key={p} label={`${lbl} [array]`} icon={<IconFolder />}>
					{val.length === 0 ? (
						<div style={{ padding: 6, color: '#6b7280' }}>Empty array</div>
					) : (
						val.map((it, i) => makeNode(`${i + 1}: ${shortPreview(it)}`, it, `${p}.${i}`))
					)}
				</TreeNode>
			);
		}

		const entries = Object.entries(val);
		return (
			<TreeNode key={p} label={`${lbl} {object}`} icon={<IconFolder />}>
				{entries.length === 0
					? <div style={{ padding: 6, color: '#6b7280' }}>Empty object</div>
					: entries.map(([k, v]) => makeNode(k, v, `${p}.${k}`))}
			</TreeNode>
		);
	};

	return makeNode(label, value, path || label);
}
