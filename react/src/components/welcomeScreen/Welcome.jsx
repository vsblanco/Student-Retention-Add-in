import React from "react";

const styles = {
	container: {
		position: "fixed",
		inset: 0,
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		background: "#ffffff",
		zIndex: 9999,
		padding: 20,
	},
	card: {
		background: "#fff",
		borderRadius: 10,
		padding: "28px 28px",
		maxWidth: 720,
		width: "100%",
		boxShadow: "0 8px 30px rgba(0,0,0,0.12)",
        border: "1px solid #eee",
	},
	header: { margin: 0, fontSize: 24, fontWeight: 600, color: "#222" },
	sub: { marginTop: 8, marginBottom: 16, color: "#555" },
	buttonRow: { display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 8 },
	primaryButton: {
		background: "#0b5cff",
		color: "#fff",
		border: "none",
		padding: "8px 14px",
		borderRadius: 6,
		cursor: "pointer",
	},
};

export default function Welcome({ onClose = () => {}, user = "" }) {
	return (
		<div style={styles.container} role="dialog" aria-modal="true" aria-label="Welcome">
			<div style={styles.card}>
				<h2 style={styles.header}>
					{user ? `Welcome, ${user}!` : "Welcome!"}
				</h2>

				<p style={styles.sub}>
					Thank you for using the Student Retention Kit.
				</p>

				<div style={styles.buttonRow}>
					<button
						type="button"
						style={styles.primaryButton}
						onClick={() => onClose()}
					>
						Continue
					</button>
				</div>
			</div>
		</div>
	);
}