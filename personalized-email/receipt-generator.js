// V-1.3 - 2025-10-01 - 6:04 PM EDT
import { getTodaysLdaSheetName } from './utils.js';

/**
 * Renders an HTML string with basic formatting into a jsPDF document,
 * respecting a maximum height for the content area.
 * Supports <p>, <strong>, <em>, <ul>, <ol>, and <li> tags.
 * @param {jsPDF} doc The jsPDF document instance.
 * @param {string} html The HTML string to render.
 * @param {object} options Configuration options (startX, startY, maxWidth, maxHeight).
 * @returns {number} The final Y position after rendering.
 */
function renderHtmlInPdf(doc, html, options) {
    let { startX, startY, maxWidth, maxHeight } = options;
    let currentY = startY;
    let isTruncated = false;

    const tempDiv = document.createElement('div');
    tempDiv.style.display = 'none';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    const processNode = (node, currentX, styles) => {
        if (isTruncated) return currentX;

        // Check if we're about to overflow before processing a new line/element
        if (currentY > startY + maxHeight - 12) { // 12 is approx line height
            if (!isTruncated) {
                doc.setFont(undefined, 'italic');
                doc.setTextColor(150); // Muted grey color for the truncation message
                doc.text("[... content truncated ...]", startX, currentY);
                isTruncated = true;
            }
            return currentX;
        }

        // Apply styles based on tags
        const isBold = styles.isBold || node.tagName === 'STRONG' || node.tagName === 'B';
        const isItalic = styles.isItalic || node.tagName === 'EM' || node.tagName === 'I';
        let fontStyle = 'normal';
        if (isBold && isItalic) fontStyle = 'bolditalic';
        else if (isBold) fontStyle = 'bold';
        else if (isItalic) fontStyle = 'italic';
        doc.setFont(undefined, fontStyle);
        doc.setTextColor(0); // Reset text color

        if (node.nodeType === 3) { // Text node
            let textContent = (node.textContent || '').replace(/\s+/g, ' ');
            const words = textContent.split(' ');
            for (const word of words) {
                if (!word) continue;
                const wordWithSpace = word + ' ';
                const wordWidth = doc.getTextWidth(wordWithSpace);
                
                if (currentX + wordWidth > startX + maxWidth) {
                    currentY += 12;
                    currentX = startX;
                    if (currentY > startY + maxHeight - 12) {
                        if (!isTruncated) {
                            doc.setFont(undefined, 'italic');
                            doc.setTextColor(150);
                            doc.text("[... content truncated ...]", startX, currentY);
                            isTruncated = true;
                        }
                        break;
                    }
                }
                doc.text(wordWithSpace, currentX, currentY);
                currentX += wordWidth;
            }
        } else { // Element node
            for (const child of Array.from(node.childNodes)) {
                if (isTruncated) break;
                currentX = processNode(child, currentX, { isBold, isItalic });
            }
        }
        return currentX;
    };
    
    Array.from(tempDiv.children).forEach(element => {
        if (isTruncated) return;
        
        switch (element.tagName) {
            case 'P':
                processNode(element, startX, {});
                currentY += 18;
                break;
            case 'UL':
            case 'OL':
                Array.from(element.children).forEach((li, index) => {
                    if (isTruncated || currentY > startY + maxHeight - 12) {
                        if (!isTruncated) {
                            doc.setFont(undefined, 'italic');
                            doc.setTextColor(150);
                            doc.text("[... content truncated ...]", startX, currentY);
                            isTruncated = true;
                        }
                        return;
                    }
                    const bullet = (element.tagName === 'OL') ? `${index + 1}. ` : '• ';
                    doc.text(bullet, startX, currentY);
                    processNode(li, startX + 15, {}); // Process list item with indent
                    currentY += 18;
                });
                break;
            default:
                processNode(element, startX, {});
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
        const maxBodyContainerHeight = 120; // Max height for the body containers
        const textPadding = 5; // Padding inside the containers
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

        // --- "Before" Template Container ---
        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        doc.text("Template Format:", margin, currentY);
        currentY += 15;
        
        const container1StartY = currentY;
        doc.setDrawColor(220, 220, 220); // Light grey border
        doc.roundedRect(margin, container1StartY, contentWidth, maxBodyContainerHeight, 3, 3, 'S');
        renderHtmlInPdf(doc, bodyTemplate, {
            startX: margin + textPadding,
            startY: container1StartY + textPadding + 2, // Start slightly inside
            maxWidth: contentWidth - (textPadding * 2),
            maxHeight: maxBodyContainerHeight - (textPadding * 2)
        });
        currentY = container1StartY + maxBodyContainerHeight + 15;

        // --- "After" Example Container ---
        doc.setFont(undefined, 'bold');
        doc.text("Example:", margin, currentY);
        currentY += 15;

        const container2StartY = currentY;
        const randomStudentPayload = payload[Math.floor(Math.random() * payload.length)];
        doc.setDrawColor(220, 220, 220);
        doc.roundedRect(margin, container2StartY, contentWidth, maxBodyContainerHeight, 3, 3, 'S');
        renderHtmlInPdf(doc, randomStudentPayload.body, {
            startX: margin + textPadding,
            startY: container2StartY + textPadding + 2,
            maxWidth: contentWidth - (textPadding * 2),
            maxHeight: maxBodyContainerHeight - (textPadding * 2)
        });
        currentY = container2StartY + maxBodyContainerHeight + 20;

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

