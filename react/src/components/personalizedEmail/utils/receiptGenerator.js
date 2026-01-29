import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { getTodaysLdaSheetName } from './helpers';

/**
 * Renders an HTML string with basic formatting into a jsPDF document,
 * with automatic page breaks (no truncation).
 * @returns {number} The final Y position after rendering
 */
function renderHtmlInPdf(doc, html, options) {
    let { startX, startY, maxWidth, margin, pageHeight } = options;
    let currentY = startY;
    const lineHeight = 12;
    const paragraphSpacing = 18;

    const tempDiv = document.createElement('div');
    tempDiv.style.display = 'none';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    const checkPageBreak = (neededHeight = lineHeight) => {
        if (currentY + neededHeight > pageHeight - margin) {
            doc.addPage();
            currentY = margin;
        }
    };

    const processNode = (node, currentX, styles) => {
        const isBold = styles.isBold || node.tagName === 'STRONG' || node.tagName === 'B';
        const isItalic = styles.isItalic || node.tagName === 'EM' || node.tagName === 'I';
        let fontStyle = 'normal';
        if (isBold && isItalic) fontStyle = 'bolditalic';
        else if (isBold) fontStyle = 'bold';
        else if (isItalic) fontStyle = 'italic';
        doc.setFont(undefined, fontStyle);
        doc.setTextColor(0);

        if (node.nodeType === 3) {
            let textContent = (node.textContent || '').replace(/\s+/g, ' ');
            const words = textContent.split(' ');
            for (const word of words) {
                if (!word) continue;
                const wordWithSpace = word + ' ';
                const wordWidth = doc.getTextWidth(wordWithSpace);

                if (currentX + wordWidth > startX + maxWidth) {
                    currentY += lineHeight;
                    currentX = startX;
                    checkPageBreak();
                }
                doc.text(wordWithSpace, currentX, currentY);
                currentX += wordWidth;
            }
        } else {
            for (const child of Array.from(node.childNodes)) {
                currentX = processNode(child, currentX, { isBold, isItalic });
            }
        }
        return currentX;
    };

    Array.from(tempDiv.children).forEach(element => {
        checkPageBreak(paragraphSpacing);

        switch (element.tagName) {
            case 'P':
                processNode(element, startX, {});
                currentY += paragraphSpacing;
                break;
            case 'UL':
            case 'OL':
                Array.from(element.children).forEach((li, index) => {
                    checkPageBreak(paragraphSpacing);
                    const bullet = (element.tagName === 'OL') ? `${index + 1}. ` : '• ';
                    doc.text(bullet, startX, currentY);
                    processNode(li, startX + 15, {});
                    currentY += paragraphSpacing;
                });
                break;
            default:
                processNode(element, startX, {});
                currentY += paragraphSpacing;
        }
    });

    document.body.removeChild(tempDiv);
    return currentY;
}

/**
 * Estimates the height needed to render HTML content
 */
function estimateHtmlHeight(doc, html, maxWidth) {
    const lineHeight = 12;
    const paragraphSpacing = 18;
    let estimatedHeight = 0;

    const tempDiv = document.createElement('div');
    tempDiv.style.display = 'none';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    const estimateNodeHeight = (node, currentX) => {
        if (node.nodeType === 3) {
            let textContent = (node.textContent || '').replace(/\s+/g, ' ');
            const words = textContent.split(' ');
            for (const word of words) {
                if (!word) continue;
                const wordWithSpace = word + ' ';
                const wordWidth = doc.getTextWidth(wordWithSpace);
                if (currentX + wordWidth > maxWidth) {
                    estimatedHeight += lineHeight;
                    currentX = 0;
                }
                currentX += wordWidth;
            }
        } else {
            for (const child of Array.from(node.childNodes)) {
                currentX = estimateNodeHeight(child, currentX);
            }
        }
        return currentX;
    };

    Array.from(tempDiv.children).forEach(element => {
        estimateNodeHeight(element, 0);
        estimatedHeight += paragraphSpacing;
    });

    document.body.removeChild(tempDiv);
    return estimatedHeight + 20; // Add some padding
}

/**
 * Generates a PDF receipt from the email payload using jsPDF and jsPDF-AutoTable.
 * @param {Array} emails - Array of email objects
 * @param {string} bodyTemplate - The email body template
 * @param {Object} initiator - Object with name and email of who initiated the send
 * @param {boolean} returnBase64 - If true, returns base64 string instead of saving
 * @returns {string|undefined} - Base64 string if returnBase64 is true, undefined otherwise
 */
export function generatePdfReceipt(emails, bodyTemplate, initiator = {}, returnBase64 = false) {
    if (!emails || emails.length === 0) {
        console.error("Emails array is empty. Cannot generate PDF receipt.");
        return;
    }

    try {
        const doc = new jsPDF({ orientation: "portrait", unit: "px", format: "letter" });

        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 30;
        const contentWidth = pageWidth - (margin * 2);
        let currentY = 0;

        // Header
        doc.setFontSize(18);
        doc.text("Email Sending Receipt", pageWidth / 2, currentY + 40, { align: "center" });
        doc.setFontSize(10);
        doc.text(`Sent on: ${new Date().toLocaleString()}`, pageWidth / 2, currentY + 55, { align: "center" });

        // Add initiator info
        if (initiator.name || initiator.email) {
            doc.text(`Initiated by: ${initiator.name || 'Unknown'}${initiator.email ? ` (${initiator.email})` : ''}`, pageWidth / 2, currentY + 68, { align: "center" });
            currentY = 88;
        } else {
            currentY = 75;
        }

        // Summary section
        doc.setFontSize(12);
        doc.text("Summary", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 15;

        doc.setFontSize(10);
        doc.text(`Total Emails Sent: ${emails.length}`, margin, currentY);
        currentY += 12;

        const senderCounts = emails.reduce((acc, email) => {
            const from = email.from || "N/A";
            acc[from] = (acc[from] || 0) + 1;
            return acc;
        }, {});

        const uniqueSenders = Object.keys(senderCounts);

        if (uniqueSenders.length === 1) {
            doc.text(`Sent From: ${uniqueSenders[0]}`, margin, currentY);
            currentY += 12;
        } else {
            doc.setFont(undefined, 'bold');
            doc.text(`Sent From (Breakdown):`, margin, currentY);
            doc.setFont(undefined, 'normal');
            currentY += 12;

            uniqueSenders.forEach(sender => {
                const count = senderCounts[sender];
                doc.text(`- ${sender}: ${count} email(s)`, margin + 10, currentY);
                currentY += 12;
            });
        }
        currentY += 20;

        // Message Body section
        doc.setFontSize(12);
        doc.text("Message Body", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 20;

        const containsParameters = /\{(\w+)\}/.test(bodyTemplate);

        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        const beforeTitle = containsParameters ? "Template Format:" : "Email Body:";
        doc.text(beforeTitle, margin, currentY);
        doc.setFont(undefined, 'normal');
        currentY += 15;

        // Render template body (full content, no truncation)
        currentY = renderHtmlInPdf(doc, bodyTemplate, {
            startX: margin,
            startY: currentY,
            maxWidth: contentWidth,
            margin: margin,
            pageHeight: pageHeight
        });

        currentY += 10;

        // Example section (if template has parameters)
        if (containsParameters) {
            const randomStudentPayload = emails[Math.floor(Math.random() * emails.length)];

            // Estimate height needed for example section
            const estimatedExampleHeight = estimateHtmlHeight(doc, randomStudentPayload.body, contentWidth);
            const spaceRemaining = pageHeight - margin - currentY;

            // If example won't fit on current page, start new page
            if (estimatedExampleHeight > spaceRemaining) {
                doc.addPage();
                currentY = margin;
            }

            doc.setFont(undefined, 'bold');
            doc.text("Example:", margin, currentY);
            doc.setFont(undefined, 'normal');
            currentY += 15;

            // Render example body (full content, no truncation)
            currentY = renderHtmlInPdf(doc, randomStudentPayload.body, {
                startX: margin,
                startY: currentY,
                maxWidth: contentWidth,
                margin: margin,
                pageHeight: pageHeight
            });

            currentY += 10;
        }

        // Recipient list on a new page
        doc.addPage();

        // Add header for recipients page
        doc.setFontSize(12);
        doc.text("Recipient List", margin, margin);
        doc.line(margin, margin + 2, pageWidth - margin, margin + 2);

        const tableColumn = ["#", "Recipient Email", "Subject"];
        const tableRows = emails.map((email, index) => [
            index + 1,
            email.to,
            email.subject.substring(0, 45) + (email.subject.length > 45 ? '...' : '')
        ]);

        autoTable(doc, {
            head: [tableColumn],
            body: tableRows,
            startY: margin + 15,
            theme: 'grid',
            headStyles: { fillColor: [41, 128, 185] },
            styles: { fontSize: 8, cellPadding: 2 },
            columnStyles: {
                0: { cellWidth: 'auto' },
                1: { cellWidth: 150 },
                2: { cellWidth: 'auto' }
            }
        });

        if (returnBase64) {
            // Return as base64 string (without the data:application/pdf;base64, prefix)
            return doc.output('datauristring').split(',')[1];
        } else {
            const fileName = `Email_Receipt_${getTodaysLdaSheetName().replace("LDA ", "")}.pdf`;
            doc.save(fileName);
        }

    } catch (error) {
        console.error("Failed to generate PDF receipt:", error);
        return undefined;
    }
}
