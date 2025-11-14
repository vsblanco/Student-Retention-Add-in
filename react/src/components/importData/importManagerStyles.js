const styles = {
	container: {
		display: 'flex',
		flexDirection: 'column',
		gap: 12,
		padding: 16,
		background: '#f7fafc',
		borderRadius: 8,
		border: '1px solid #e2e8f0',
		maxWidth: 720,
		width: '100%',
		boxSizing: 'border-box',
		fontFamily: 'Inter, system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial',
	},
	controlRow: {
		display: 'flex',
		gap: 8,
		alignItems: 'center',
		flexWrap: 'wrap',
	},
	uploadButton: {
		background: '#2563eb',
		color: '#fff',
		border: 'none',
		padding: '8px 12px',
		borderRadius: 6,
		cursor: 'pointer',
		boxShadow: '0 1px 3px rgba(15,23,42,0.06)',
	},
	importButton: {
		background: '#454c57',
		color: '#fff',
		border: 'none',
		padding: '8px 12px',
		borderRadius: 6,
		cursor: 'pointer',
		boxShadow: '0 1px 3px rgba(16,185,129,0.08)',
	},
	infoBox: {
		background: '#fff',
		padding: '10px 12px',
		borderRadius: 6,
		border: '1px solid #e6edf3',
		boxShadow: '0 1px 2px rgba(2,6,23,0.04)',
	},
	fileName: {
		fontWeight: 600,
		color: '#0f172a',
	},
	statusText: {
		color: '#475569',
		fontSize: 13,
	},
	processorWrap: {
		marginTop: 8,
		padding: 12,
		borderRadius: 6,
		background: '#ffffff',
		border: '1px dashed #c7d2fe',
	},
};

export default styles;
