import React from 'react';

export default function DeleteConfirmModal({ isOpen, title = 'Confirm delete', message = '', confirmLabel = 'Delete', onConfirm = () => {}, onCancel = () => {} }) {
	if (!isOpen) return null;
	return (
		<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.35)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10000 }}>
			<div style={{ width: 'min(520px, 94%)', background: '#fff', borderRadius: 8, padding: 16, boxShadow: '0 8px 32px rgba(2,6,23,0.14)', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>
				<h4 style={{ margin: '0 0 8px 0' }}>{title}</h4>
				<div style={{ color: '#374151', marginBottom: 12 }}>{message}</div>
				<div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
					<button onClick={onCancel} style={{ padding: '8px 10px', borderRadius: 8, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}>Cancel</button>
					<button onClick={onConfirm} style={{ padding: '8px 10px', borderRadius: 8, background: '#ef4444', color: '#fff', border: 'none', cursor: 'pointer' }}>{confirmLabel}</button>
				</div>
			</div>
		</div>
	);
}
