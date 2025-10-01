// V-1.0 - 2025-10-01 - 3:44 PM EDT
import { getTodaysLdaSheetName } from './utils.js';

/**
 * Generates a PDF receipt from the email payload using jsPDF and jsPDF-AutoTable.
 * @param {Array<object>} payload - The array of email objects that were sent.
 */
export function generatePdfReceipt(payload) {
    if (!payload || payload.length === 0) {
        console.error("Payload is empty. Cannot generate PDF receipt.");
        return;
    }

    try {
        // Initialize jsPDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        // --- PDF CONTENT ---

        // 1. Header
        doc.setFontSize(22);
        doc.text("Email Sending Receipt", 105, 20, null, null, "center");
        doc.setFontSize(12);
        const generationDate = new Date().toLocaleString();
        doc.text(`Generated on: ${generationDate}`, 105, 30, null, null, "center");

        // 2. Summary
        doc.setFontSize(14);
        doc.text("Summary", 14, 45);
        doc.setLineWidth(0.5);
        doc.line(14, 46, 196, 46); // Underline

        doc.setFontSize(12);
        doc.text(`Total Emails Sent: ${payload.length}`, 14, 55);
        // Get the first "from" address as a representative sender
        const sender = payload[0]?.from || "N/A";
        doc.text(`Sent From: ${sender}`, 14, 62);

        // 3. Table of Recipients
        const tableColumn = ["#", "Recipient Email", "Subject"];
        const tableRows = [];

        payload.forEach((email, index) => {
            const emailData = [
                index + 1,
                email.to,
                // Truncate long subjects to prevent table overflow
                email.subject.substring(0, 45) + (email.subject.length > 45 ? '...' : '') 
            ];
            tableRows.push(emailData);
        });

        // Add table using autoTable plugin, which handles pagination automatically
        doc.autoTable({
            head: [tableColumn],
            body: tableRows,
            startY: 75,
            theme: 'grid',
            headStyles: { fillColor: [41, 128, 185] }, // A blue header color
            styles: { fontSize: 9 },
            columnStyles: {
                0: { cellWidth: 10 },
                1: { cellWidth: 75 },
                2: { cellWidth: 'auto' }
            }
        });

        // --- PDF SAVING ---
        const fileName = `Email_Receipt_${getTodaysLdaSheetName().replace("LDA ", "")}.pdf`;
        doc.save(fileName);

    } catch (error) {
        console.error("Failed to generate PDF receipt:", error);
        // You could show an error message to the user here if desired
    }
}
