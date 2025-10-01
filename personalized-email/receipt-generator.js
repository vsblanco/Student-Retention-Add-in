// V-1.2 - 2025-10-01 - 5:15 PM EDT
import { getTodaysLdaSheetName } from './utils.js';

/**
 * Renders an HTML string with basic formatting into a jsPDF document.
 * Supports <p>, <strong>, <em>, <ul>, <ol>, and <li> tags.
 * @param {jsPDF} doc The jsPDF document instance.
 * @param {string} html The HTML string to render.
 * @param {object} options Configuration options (startX, startY, maxWidth).
 * @returns {number} The final Y position after rendering.
 */
function renderHtmlInPdf(doc, html, options) {
    let { startX, startY, maxWidth } = options;
    let currentY = startY;

    const tempDiv = document.createElement('div');
    tempDiv.style.display = 'none';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    const processNode = (node, currentX, inheritedStyles = {}) => {
        let textContent = (node.textContent || '').replace(/\s+/g, ' ');

        // Apply styles
        let isBold = inheritedStyles.isBold || node.tagName === 'STRONG' || node.tagName === 'B';
        let isItalic = inheritedStyles.isItalic || node.tagName === 'EM' || node.tagName === 'I';
        let fontStyle = 'normal';
        if (isBold && isItalic) fontStyle = 'bolditalic';
        else if (isBold) fontStyle = 'bold';
        else if (isItalic) fontStyle = 'italic';
        doc.setFont(undefined, fontStyle);

        if (node.childNodes.length > 0 && node.tagName !== 'STRONG' && node.tagName !== 'EM' && node.tagName !== 'B' && node.tagName !== 'I') {
            let newX = currentX;
            for (const child of Array.from(node.childNodes)) {
                newX = processNode(child, newX, { isBold, isItalic });
            }
            return newX;
        } else {
             const words = textContent.split(' ');
             for (const word of words) {
                if (!word) continue;
                const wordWidth = doc.getTextWidth(word + ' ');
                if (currentX + wordWidth > startX + maxWidth) {
                    currentY += 12; // Line height
                    currentX = startX;
                }
                doc.text(word, currentX, currentY);
                currentX += wordWidth;
             }
             return currentX;
        }
    };
    
    Array.from(tempDiv.children).forEach(element => {
        if (currentY > doc.internal.pageSize.height - 40) doc.addPage();
        
        switch (element.tagName) {
            case 'P':
                processNode(element, startX);
                currentY += 18;
                break;
            case 'UL':
            case 'OL':
                Array.from(element.children).forEach((li, index) => {
                    const bullet = (element.tagName === 'OL') ? `${index + 1}. ` : '• ';
                    doc.text(bullet, startX, currentY);
                    const originalStartX = startX;
                    startX += 15; // Indent
                    maxWidth -= 15;
                    processNode(li, startX);
                    currentY += 18;
                    startX = originalStartX; // Reset indent
                    maxWidth += 15;
                });
                break;
            default:
                processNode(element, startX);
                currentY += 18;
        }
    });

    document.body.removeChild(tempDiv);
    return currentY;
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
        const doc = new jsPDF({ orientation: "portrait", unit: "px", format: "letter" });
        
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 30;
        const contentWidth = pageWidth - (margin * 2);
        let currentY = 0;

        // --- PDF CONTENT ---
        doc.setFontSize(18);
        doc.text("Email Sending Receipt", pageWidth / 2, currentY + 40, { align: "center" });
        doc.setFontSize(10);
        doc.text(`Generated on: ${new Date().toLocaleString()}`, pageWidth / 2, currentY + 55, { align: "center" });
        currentY = 75;

        doc.setFontSize(12);
        doc.text("Summary", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 15;

        doc.setFontSize(10);
        doc.text(`Total Emails Sent: ${payload.length}`, margin, currentY);
        doc.text(`Sent From: ${payload[0]?.from || "N/A"}`, margin, currentY + 12);
        currentY += 32;

        // --- Message Body Section ---
        doc.setFontSize(12);
        doc.text("Message Body", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 20;

        // --- "Before" Template ---
        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        doc.text("Template Format (Before Personalization):", margin, currentY);
        currentY += 15;
        currentY = renderHtmlInPdf(doc, bodyTemplate, { startX: margin, startY: currentY, maxWidth: contentWidth });
        currentY += 10;

        // --- "After" Example ---
        doc.setFont(undefined, 'bold');
        doc.text("Example (After Personalization for a Random Student):", margin, currentY);
        currentY += 15;
        const randomStudentPayload = payload[Math.floor(Math.random() * payload.length)];
        currentY = renderHtmlInPdf(doc, randomStudentPayload.body, { startX: margin, startY: currentY, maxWidth: contentWidth });
        currentY += 20;

        // --- Table of Recipients ---
        const tableColumn = ["#", "Recipient Email", "Subject"];
        const tableRows = payload.map((email, index) => [
            index + 1,
            email.to,
            email.subject.substring(0, 45) + (email.subject.length > 45 ? '...' : '')
        ]);

        doc.autoTable({
            head: [tableColumn],
            body: tableRows,
            startY: currentY,
            theme: 'grid',
            headStyles: { fillColor: [41, 128, 185] },
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

