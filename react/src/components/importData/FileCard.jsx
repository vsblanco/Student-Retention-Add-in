import React from 'react';
import csvIcon from '../../assets/icons/csv-icon.png';
import { File } from 'lucide-react';

export default function FileCard({ file, rows, type, action, icon } = {}) {
	const name = (file && (file.name || file.filename)) || 'Unknown.csv';
	const sizeKB = file && file.size ? Math.round(file.size / 1024) : null;

	// small badge style for type/action
	const badgeStyle = {
		background: '#eef2ff',
		color: '#2b6cb0',
		fontSize: 11,
		fontWeight: 700,
		padding: '4px 4px',
		borderRadius: 999,
		whiteSpace: 'nowrap',
	};

	return (
		<div
			style={{
				width: '100%',
				boxSizing: 'border-box',
				display: 'flex',
				alignItems: 'center',
				gap: 7,
				padding: '10px 12px',
				borderRadius: 8,
				border: '1px solid rgba(15,23,42,0.06)',
				background: '#fff',
				boxShadow: '0 1px 2px rgba(16,24,40,0.03)',
			}}
			title={name}
		>
			<div style={{ width: 44, height: 44, flex: '0 0 44px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
				{ /* if an import-type icon was provided, show it (overrides csvIcon) */ }
				{icon ? (
					<img src={icon} alt="import type" style={{ width: 36, height: 36, objectFit: 'contain' }} />
				) : /\.csv$/i.test(name) ? (
					<img src={csvIcon} alt="CSV" style={{ width: 36, height: 36, objectFit: 'contain' }} />
				) : (
					<File size={36} color="#9aa4b2" />
				)}
			</div>

			<div style={{ flex: 1, minWidth: 0 }}>
				<div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
					{name}
				</div>

				<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
					<div style={{ fontSize: 12, color: '#556070' }}>
						{sizeKB !== null ? `${sizeKB} KB` : '—'}{rows !== undefined ? ` • ${rows} rows` : ''}
					</div>

					{/* show type/action badges if provided */}
					{(type || action) && (
						<div style={{ display: 'flex', gap: 1, marginLeft: 4 }}>
							{type && <div style={badgeStyle}>{type}</div>}
							{action && <div style={{ ...badgeStyle, background: '#fff7ed', color: '#d97706' }}>{action}</div>}
						</div>
					)}
				</div>
			</div>
		</div>
	);
}
