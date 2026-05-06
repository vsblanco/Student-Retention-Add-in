import React, { useState, useEffect } from "react";
import Media from "./Media";
import chromeExtensionService from "../../services/chromeExtensionService.js";
import importdatataskpane from "../../assets/tutorial/importdatataskpane.png";
import createlda from "../../assets/tutorial/createlda.gif";

export default function Tutorial({ pages = null, onBack = () => {}, onClose = () => {}, onFinish = null }) {
    
    // !!! UPDATE THIS URL TO YOUR ACTUAL CHROME WEB STORE LINK !!!
    const CHROME_EXTENSION_URL = "https://chrome.google.com/webstore/detail/your-extension-id";

    // 1. Define default pages data
    const defaultPagesData = [
        { 
            title: "What is Student Retention Kit?", 
            content: <p>The Student Retention Kit is a tool designed to help educators identify and support at-risk students. It's goal is to make your workflow as effiecently as possible. So that you can focus on what's most important.</p> 
        },
        {
            title: "What can I do with this?",
            content: <p>There's a variety of features bundled in this kit. They include:</p>,
            bullets: [
                "Importing external reports onto your sheets",
                "Automatic LDA creation",
                "Sending Personalized emails to students",
                "Real time student submission feedback",
                'Student communication tracking',
            ],
        },
        {
            title: "Initial Setup",
            content: <p>Before we continue further, let's make sure your workbook is set up correctly. You can skip this however, your features may be limitted.</p>,
            checklist: [
                { label: "Master List Sheet", status: false, createSheet: "Master List" },
                { label: "Student History Sheet", status: false, createSheet: "Student History" },
                { label: "Missing Assignments Sheet", status: false, createSheet: "Missing Assignments" },
                { label: "Student Retention Kit - Chrome Extension", status: false, id: "extension-check" },
            ],
        },
        {
            title: "Master List",
           content: <p>This sheet contains the complete student population of your campus. It serves as the target for your imports and the source for your LDA.</p>,
        },
         {
            title: "Student History",
           content: <p>This sheet maintains a record of student interactions and communications. New entries can be added through the Student View panel. There are two methods for making a comment: you can either manually create a comment from the Student View panel or type in the Outreach Column, which will automatically generate one.</p>,
        },
         {
            title: "Missing Assignments",
           content: <p>This sheet contains a list of students' missing assignments. The report is generated using the Student Retention Kit—Chrome Extension.</p>,
        },
        {
            title: "Importing Data",
           content: <p>Importing Data is handled through the Student Retention Kit—Chrome Extension. All it needs is a student population report with either a SyStudentId or Student SIS column in it. It will automatically organize the data and create a canvas for their grade book information. The system will automatically import the data into the Master List sheet for your viewing pleasure.</p>,
        },
           {
            title: "Creating the LDA sheet.",
            component: (
                <>
                    <div style={{ padding: 10 }}>
                        <p>You can create the LDA sheet by clicking on the Create LDA button on the ribbon.</p>
                    </div>
                    <Media src={createlda} alt="Import Data taskpane" width="820px" fit="contain" clickable={false} />
                 <div style={{ padding: 10 }}>
                        <p>Imports will appear on the Master List sheet.</p>
                    </div>
                </>
                
            ),
        },
        {
            title: "Sending Personalized Emails",
           content: <p>You can automatically send personalized emails to each student by granting consent to send emails on your behalf. You will see what parameters you can choose from to help personalize your emails. You can also create your own parameters for further customization. Alternatively, if you have a power automate premium license, you can configure an HTTP request to create more advanced automations.</p>,
        },
        {
            title: "That's it!",
           content: <p>The mission of this application is to make your job more efficient. So you have more time to focus on what's most important. The students.</p>,
        },
    ];

    const [tutorialPages, setTutorialPages] = useState(pages && pages.length ? pages : defaultPagesData);
    const [index, setIndex] = useState(0);

    // 2. Main Logic: Check for Sheets and Extension on mount
    useEffect(() => {
        const checkExistingSheets = async () => {
            try {
                await Excel.run(async (context) => {
                    const sheets = context.workbook.worksheets;
                    const sheetsToCheck = ["Master List", "Student History", "Missing Assignments"];

                    const sheetProxies = sheetsToCheck.map(name => ({
                        name: name,
                        proxy: sheets.getItemOrNullObject(name)
                    }));

                    await context.sync();

                    const existingMap = {};
                    sheetProxies.forEach(item => { existingMap[item.name] = !item.proxy.isNullObject; });

                    setTutorialPages(prev => {
                        const newPages = [...prev];
                        const setupIndex = newPages.findIndex(p => p.checklist);
                        if (setupIndex !== -1) {
                            const setupPage = { ...newPages[setupIndex] };
                            setupPage.checklist = setupPage.checklist.map(item => {
                                if (item.createSheet && existingMap[item.createSheet]) return { ...item, status: true };
                                return item;
                            });
                            newPages[setupIndex] = setupPage;
                        }
                        return newPages;
                    });
                });
            } catch (e) { console.error(e); }
        };

        const checkExtension = () => {
            // Check if extension is already detected (via App.jsx master relay)
            if (chromeExtensionService.getInstallationStatus()) {
                console.log("Tutorial: Extension already detected by master relay");
                setTutorialPages(prev => {
                    const newPages = [...prev];
                    const setupIndex = newPages.findIndex(p => p.checklist);
                    if (setupIndex !== -1) {
                        const setupPage = { ...newPages[setupIndex] };
                        setupPage.checklist = setupPage.checklist.map(item => {
                            if (item.id === "extension-check") return { ...item, status: true };
                            return item;
                        });
                        newPages[setupIndex] = setupPage;
                    }
                    return newPages;
                });
            }

            // Listen for extension installation events
            const removeListener = chromeExtensionService.addListener((event) => {
                if (event.type === "installed") {
                    console.log("Tutorial: Extension installed event received from service");
                    setTutorialPages(prev => {
                        const newPages = [...prev];
                        const setupIndex = newPages.findIndex(p => p.checklist);
                        if (setupIndex !== -1) {
                            const setupPage = { ...newPages[setupIndex] };
                            setupPage.checklist = setupPage.checklist.map(item => {
                                if (item.id === "extension-check") return { ...item, status: true };
                                return item;
                            });
                            newPages[setupIndex] = setupPage;
                        }
                        return newPages;
                    });
                }
            });

            return removeListener;
        };

        if (!pages) {
            checkExistingSheets();
            const cleanupExtension = checkExtension();
            return cleanupExtension;
        }
    }, []); 

    const prev = () => setIndex(i => Math.max(0, i - 1));
    const next = () => setIndex(i => Math.min(tutorialPages.length - 1, i + 1));
    const goTo = (i) => setIndex(Math.max(0, Math.min(tutorialPages.length - 1, i)));
    
    // Updated finish handler
    const finish = () => {
        if (onFinish) {
            onFinish();
        } else {
            onClose();
        }
    };

    const handleCreateSheet = async (sheetName, pageIndex, itemIndex) => {
        try {
            await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                const existingSheet = sheets.getItemOrNullObject(sheetName);
                await context.sync();
                if (existingSheet.isNullObject) sheets.add(sheetName);
                await context.sync();
            });
            setTutorialPages(prev => {
                const newPages = [...prev];
                newPages[pageIndex].checklist[itemIndex].status = true;
                return newPages;
            });
        } catch (error) {
            if (error instanceof OfficeExtension.Error) {
                 setTutorialPages(prev => {
                    const newPages = [...prev];
                    newPages[pageIndex].checklist[itemIndex].status = true;
                    return newPages;
                });
            }
        }
    };

    // Styles (kept same as provided)
    const styles = {
        card: { position: "absolute", inset: 0, background: "#fff", borderRadius: 0, padding: 0, width: "100%", height: "100%", maxWidth: "none", boxShadow: "none" },
        header: { margin: 0, fontSize: 22, fontWeight: 600, color: "#222", display: "flex", justifyContent: "space-between", alignItems: "center" },
        sub: { marginTop: 8, marginBottom: 12, color: "#555" },
        content: { minHeight: 140, marginTop: 12 },
        bulletList: { margin: "10px 0 0 0", padding: 0, listStyle: "none", color: "#333" },
        bulletItem: { display: "flex", gap: 10, alignItems: "flex-start", padding: "6px 0" },
        bulletIcon: { flex: "0 0 18px", marginTop: 4 },
        checklistList: { margin: "12px 0 0 0", padding: 0, listStyle: "none" },
        checklistItem: { display: "flex", gap: 10, alignItems: "center", padding: "6px 0", color: "#333" },
        checklistIcon: { flex: "0 0 18px" },
        dots: { display: "flex", gap: 8, alignItems: "center", marginTop: 12 },
        dot: (active) => ({ width: 10, height: 10, borderRadius: "50%", background: active ? "#0b5cff" : "#d7dbe9", cursor: "pointer" }),
        buttonRow: { display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 12 },
        primaryButton: { background: "#0b5cff", color: "#fff", border: "none", padding: "8px 14px", borderRadius: 6, cursor: "pointer" },
        tertiary: { background: "transparent", color: "#666", border: "none", padding: "6px 10px", cursor: "pointer" },
        createButton: { background: "#fff", border: "1px solid #0b5cff", color: "#0b5cff", padding: "4px 8px", fontSize: "12px", borderRadius: 4, cursor: "pointer", marginLeft: "auto" },
        createButtonDisabled: { background: "#f3f3f3", border: "1px solid #ddd", color: "#888", padding: "4px 8px", fontSize: "12px", borderRadius: 4, cursor: "default", marginLeft: "auto" },
    };

    const current = tutorialPages[index];
    const isLastPage = index === tutorialPages.length - 1;

    return (
        <div style={styles.card} role="dialog" aria-modal="true" aria-label="Tutorial">
            <div style={styles.header}>
                <span>{current.title || `Page ${index + 1}`}</span>
                <span style={{ fontSize: 13, color: "#777" }}>{`${index + 1} / ${tutorialPages.length}`}</span>
            </div>

            <div style={styles.content}>
                {current.component ? current.component : current.content}

                {current.bullets && (
                    <ul style={styles.bulletList}>
                        {current.bullets.map((b, i) => (
                            <li key={i} style={styles.bulletItem}>
                                <svg viewBox="0 0 24 24" width="18" height="18" style={styles.bulletIcon}><circle cx="12" cy="12" r="6" fill="#4f4e4e" /></svg>
                                <span>{b}</span>
                            </li>
                        ))}
                    </ul>
                )}

                {current.checklist && (
                    <ul style={styles.checklistList}>
                        {current.checklist.map((c, i) => {
                            const label = typeof c === "string" ? c : c.label;
                            const done = typeof c === "string" ? false : !!c.status;
                            const sheetToCreate = c.createSheet;
                            const isExtension = c.id === "extension-check";

                            return (
                                <li key={i} style={styles.checklistItem}>
                                    <span style={styles.checklistIcon}>
                                        {done ? (
                                            <svg width="18" height="18" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#28a745" /><path d="M7 12.5l2.5 2.5L17 8" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" fill="none" /></svg>
                                        ) : (
                                            <svg width="18" height="18" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9" fill="#e6e9ef" /></svg>
                                        )}
                                    </span>
                                    <span>{label}</span>

                                    {sheetToCreate && (
                                        <button 
                                            style={done ? styles.createButtonDisabled : styles.createButton}
                                            disabled={done}
                                            onClick={() => handleCreateSheet(sheetToCreate, index, i)}
                                        >
                                            {done ? "Created" : "Create"}
                                        </button>
                                    )}

                                    {isExtension && (
                                        <button 
                                            style={done ? styles.createButtonDisabled : styles.createButton}
                                            disabled={done}
                                            onClick={() => window.open(CHROME_EXTENSION_URL, "_blank")}
                                        >
                                            {done ? "Installed" : "Download"}
                                        </button>
                                    )}
                                </li>
                            );
                        })}
                    </ul>
                )}
            </div>

            <div style={styles.dots}>
                {tutorialPages.map((_, i) => (
                    <div key={i} style={styles.dot(i === index)} onClick={() => goTo(i)} />
                ))}
            </div>

            <div style={styles.buttonRow}>
                {index > 0 && <button style={styles.tertiary} onClick={prev}>Prev</button>}
                
                {/* Logic to swap Next for Finish on the last page */}
                {isLastPage ? (
                    <button style={styles.primaryButton} onClick={finish}>Finish</button>
                ) : (
                    <button style={styles.primaryButton} onClick={next}>Next</button>
                )}
            </div>
        </div>
    );
}