import { getTodaysLdaSheetName } from './helpers';

/**
 * Renders an HTML string with basic formatting into a jsPDF document,
 * respecting a maximum height for the content area.
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

        if (currentY > startY + maxHeight - 12) {
            if (!isTruncated) {
                doc.setFont(undefined, 'italic');
                doc.setTextColor(150);
                doc.text("[... content truncated ...]", startX, currentY);
                isTruncated = true;
            }
            return currentX;
        }

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
        } else {
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
                    processNode(li, startX + 15, {});
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
        const maxBodyContainerHeight = 120;
        const textPadding = 5;
        let currentY = 0;

        doc.setFontSize(18);
        doc.text("Email Sending Receipt", pageWidth / 2, currentY + 40, { align: "center" });
        doc.setFontSize(10);
        doc.text(`Sent on: ${new Date().toLocaleString()}`, pageWidth / 2, currentY + 55, { align: "center" });
        currentY = 75;

        doc.setFontSize(12);
        doc.text("Summary", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 15;

        doc.setFontSize(10);
        doc.text(`Total Emails Sent: ${payload.length}`, margin, currentY);
        currentY += 12;

        const senderCounts = payload.reduce((acc, email) => {
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

        doc.setFontSize(12);
        doc.text("Message Body", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 20;

        const containsParameters = /\{(\w+)\}/.test(bodyTemplate);

        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        const beforeTitle = containsParameters ? "Template Format:" : "Email Body:";
        doc.text(beforeTitle, margin, currentY);
        currentY += 15;

        const container1StartY = currentY;
        doc.setDrawColor(220, 220, 220);
        doc.roundedRect(margin, container1StartY, contentWidth, maxBodyContainerHeight, 3, 3, 'S');
        renderHtmlInPdf(doc, bodyTemplate, {
            startX: margin + textPadding,
            startY: container1StartY + textPadding + 2,
            maxWidth: contentWidth - (textPadding * 2),
            maxHeight: maxBodyContainerHeight - (textPadding * 2)
        });
        currentY = container1StartY + maxBodyContainerHeight + 15;

        if (containsParameters) {
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
        } else {
            currentY += 5;
        }

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

        const fileName = `Email_Receipt_${getTodaysLdaSheetName().replace("LDA ", "")}.pdf`;
        doc.save(fileName);

    } catch (error) {
        console.error("Failed to generate PDF receipt:", error);
    }
}
