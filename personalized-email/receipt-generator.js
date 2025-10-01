// V-1.1 - 2025-10-01 - 4:56 PM EDT
import { getTodaysLdaSheetName } from './utils.js';

/**
 * Helper function to strip HTML tags for clean text rendering in the PDF.
 * @param {string} html - The HTML string to clean.
 * @returns {string} The plain text content.
 */
function stripHtml(html) {
    try {
        // Use DOMParser to convert HTML string to a document, then extract text content.
        const doc = new DOMParser().parseFromString(html, 'text/html');
        return doc.body.textContent || "";
    } catch (e) {
        console.error("Could not parse HTML", e);
        return html; // Fallback to original html if parsing fails
    }
}

/**
 * Generates a PDF receipt from the email payload using jsPDF and jsPDF-AutoTable.
 * @param {Array<object>} payload - The array of email objects that were sent.
 * @param {string} bodyTemplate - The raw HTML string of the body from the editor.
 */
export function generatePdfReceipt(payload, bodyTemplate) {
    if (!payload || payload.length === 0) {
        console.error("Payload is empty. Cannot generate PDF receipt.");
        return;
    }

    try {
        const { jsPDF } = window.jspdf;
        // Initialize jsPDF with portrait, pixel units, and letter size for consistency.
        const doc = new jsPDF({
            orientation: "portrait",
            unit: "px",
            format: "letter"
        });
        
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 30;
        let currentY = 0; // Keep track of the vertical position on the page

        // --- PDF CONTENT ---

        // 1. Header
        doc.setFontSize(18);
        doc.text("Email Sending Receipt", pageWidth / 2, currentY + 40, { align: "center" });
        doc.setFontSize(10);
        const generationDate = new Date().toLocaleString();
        doc.text(`Generated on: ${generationDate}`, pageWidth / 2, currentY + 55, { align: "center" });
        currentY = 75;

        // 2. Summary
        doc.setFontSize(12);
        doc.text("Summary", margin, currentY);
        doc.setLineWidth(0.5);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2); // Underline
        currentY += 15;

        doc.setFontSize(10);
        doc.text(`Total Emails Sent: ${payload.length}`, margin, currentY);
        const sender = payload[0]?.from || "N/A";
        doc.text(`Sent From: ${sender}`, margin, currentY + 12);
        currentY += 32;

        // 3. Body Message Section
        doc.setFontSize(12);
        doc.text("Message Body", margin, currentY);
        doc.setLineWidth(0.5);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2); // Underline
        currentY += 15;

        // --- "Before" Template ---
        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        doc.text("Template Format (Before Personalization):", margin, currentY);
        currentY += 12;
        
        doc.setFont(undefined, 'normal');
        const beforeText = stripHtml(bodyTemplate);
        const beforeLines = doc.splitTextToSize(beforeText, pageWidth - (margin * 2));
        doc.text(beforeLines, margin, currentY);
        currentY += (beforeLines.length * 10) + 15; // Calculate Y position after the 'before' text

        // --- "After" Example ---
        doc.setFont(undefined, 'bold');
        doc.text("Example (After Personalization for a Random Student):", margin, currentY);
        currentY += 12;

        // Pick a random student from the payload for the example
        const randomStudentPayload = payload[Math.floor(Math.random() * payload.length)];
        
        doc.setFont(undefined, 'normal');
        const afterText = stripHtml(randomStudentPayload.body);
        const afterLines = doc.splitTextToSize(afterText, pageWidth - (margin * 2));
        doc.text(afterLines, margin, currentY);
        currentY += (afterLines.length * 10) + 20; // Calculate final Y before the table

        // 4. Table of Recipients
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
            startY: currentY,
            theme: 'grid',
            headStyles: { fillColor: [41, 128, 185] }, // A blue header color
            styles: { fontSize: 8, cellPadding: 2 },
            columnStyles: {
                0: { cellWidth: 'auto' },
                1: { cellWidth: 150 },
                2: { cellWidth: 'auto' }
            }
        });

        // --- PDF SAVING ---
        const fileName = `Email_Receipt_${getTodaysLdaSheetName().replace("LDA ", "")}.pdf`;
        doc.save(fileName);

    } catch (error) {
        console.error("Failed to generate PDF receipt:", error);
    }
}

