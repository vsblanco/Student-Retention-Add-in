import React, { useState } from "react";
import Tutorial from "./Tutorial";

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
	bullets: { margin: "12px 0 18px 20px", color: "#444" },
	buttonRow: { display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 8 },
	primaryButton: {
		background: "#0b5cff",
		color: "#fff",
		border: "none",
		padding: "8px 14px",
		borderRadius: 6,
		cursor: "pointer",
	},
	secondaryButton: {
		background: "transparent",
		color: "#0b5cff",
		border: "1px solid #0b5cff",
		padding: "6px 12px",
		borderRadius: 6,
		cursor: "pointer",
	},
	tertiary: { background: "transparent", color: "#666", border: "none", padding: "6px 10px", cursor: "pointer" },
};

export default function Welcome({ onStart = () => {}, onClose = () => {}, user = "", docsUrl = "https://example.com/docs" }) {
	const [showTutorial, setShowTutorial] = useState(false);
	const [hasFinished, setHasFinished] = useState(false);

	return (
		<div style={styles.container} role="dialog" aria-modal="true" aria-label="Welcome">
			{showTutorial ? (
				<Tutorial
					onBack={() => {
						setShowTutorial(false);
					}}
					onClose={() => {
						onClose();
					}}
					onFinish={() => {
						setShowTutorial(false);
						setHasFinished(true);
					}}
				/>
			) : (
				<div style={styles.card}>
					<h2 style={styles.header}>
                        {hasFinished 
                            ? "You are all good to go!" 
                            : (user ? `Welcome, ${user}!` : "Welcome!")
                        }
					</h2>
					
					<p style={styles.sub}>
						{hasFinished 
							? "Your workbook has been set up successfully. You can now start using the Student Retention Kit." 
							: "Would you like a tutorial on how to use the Student Retention Kit?"
						}
					</p>

					<div style={styles.buttonRow}>
						{hasFinished ? (
							<button
								type="button"
								style={styles.primaryButton}
								onClick={() => onClose()}
							>
								Done
							</button>
						) : (
							<>
								<button type="button" style={styles.tertiary} onClick={() => onClose()}>
									Dismiss
								</button>

								<button
									type="button"
									style={styles.primaryButton}
									onClick={() => {
										setShowTutorial(true);
									}}
								>
									Next
								</button>
							</>
						)}
					</div>
				</div>
			)}
		</div>
	);
}