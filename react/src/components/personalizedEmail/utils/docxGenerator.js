import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    ExternalHyperlink,
    AlignmentType,
} from 'docx';

const BULLET_REF = 'email-bullet';
const ORDERED_REF = 'email-ordered';

const NUMBERING_CONFIG = {
    config: [
        {
            reference: BULLET_REF,
            levels: [
                { level: 0, format: 'bullet', text: '•', alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                { level: 1, format: 'bullet', text: '◦', alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
            ],
        },
        {
            reference: ORDERED_REF,
            levels: [
                { level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                { level: 1, format: 'lowerLetter', text: '%2.', alignment: AlignmentType.LEFT,
                  style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
            ],
        },
    ],
};

const HEX_RE = /^#?([0-9a-fA-F]{6})$/;
const RGB_RE = /^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/;
const ALIGNMENT_MAP = {
    left: AlignmentType.LEFT,
    right: AlignmentType.RIGHT,
    center: AlignmentType.CENTER,
    justify: AlignmentType.JUSTIFIED,
};

function parseColor(value) {
    if (!value) return undefined;
    const trimmed = String(value).trim();
    const hexMatch = trimmed.match(HEX_RE);
    if (hexMatch) return hexMatch[1].toUpperCase();
    const rgbMatch = trimmed.match(RGB_RE);
    if (rgbMatch) {
        return rgbMatch.slice(1, 4)
            .map(n => Math.max(0, Math.min(255, parseInt(n, 10))).toString(16).padStart(2, '0'))
            .join('')
            .toUpperCase();
    }
    return undefined;
}

function parseStyle(styleAttr) {
    const out = {};
    if (!styleAttr) return out;
    for (const declaration of String(styleAttr).split(';')) {
        const idx = declaration.indexOf(':');
        if (idx === -1) continue;
        const prop = declaration.slice(0, idx).trim().toLowerCase();
        const val = declaration.slice(idx + 1).trim();
        if (prop && val) out[prop] = val;
    }
    return out;
}

function inheritFormat(parent, element) {
    const fmt = { ...parent };
    const tag = element.tagName.toLowerCase();
    if (tag === 'strong' || tag === 'b') fmt.bold = true;
    if (tag === 'em' || tag === 'i') fmt.italics = true;
    if (tag === 'u') fmt.underline = true;
    if (tag === 's' || tag === 'strike' || tag === 'del') fmt.strike = true;

    const style = parseStyle(element.getAttribute('style'));
    const color = parseColor(style.color);
    if (color) fmt.color = color;
    const bg = parseColor(style['background-color'] || style.background);
    if (bg) fmt.highlight = bg;
    if (style['font-weight'] === 'bold' || parseInt(style['font-weight'], 10) >= 600) fmt.bold = true;
    if (style['font-style'] === 'italic') fmt.italics = true;
    if (style['text-decoration'] && style['text-decoration'].includes('underline')) fmt.underline = true;
    return fmt;
}

function toRunOptions(fmt) {
    const opts = {};
    if (fmt.bold) opts.bold = true;
    if (fmt.italics) opts.italics = true;
    if (fmt.underline) opts.underline = {};
    if (fmt.strike) opts.strike = true;
    if (fmt.color) opts.color = fmt.color;
    if (fmt.highlight) opts.shading = { type: 'clear', color: 'auto', fill: fmt.highlight };
    return opts;
}

function buildRunsFromInline(node, format, runs) {
    if (node.nodeType === 3) {
        const text = node.textContent;
        if (text) runs.push(new TextRun({ text, ...toRunOptions(format) }));
        return;
    }
    if (node.nodeType !== 1) return;

    const tag = node.tagName.toLowerCase();

    if (tag === 'br') {
        runs.push(new TextRun({ break: 1, ...toRunOptions(format) }));
        return;
    }

    if (tag === 'a') {
        const href = node.getAttribute('href');
        const linkFormat = inheritFormat({ ...format, underline: true, color: format.color || '0563C1' }, node);
        const linkRuns = [];
        for (const child of node.childNodes) {
            buildRunsFromInline(child, linkFormat, linkRuns);
        }
        if (href && linkRuns.length > 0) {
            runs.push(new ExternalHyperlink({ link: href, children: linkRuns }));
        } else {
            runs.push(...linkRuns);
        }
        return;
    }

    const childFormat = inheritFormat(format, node);
    for (const child of node.childNodes) {
        buildRunsFromInline(child, childFormat, runs);
    }
}

function collectRuns(blockEl, baseFormat) {
    const runs = [];
    const fmt = inheritFormat(baseFormat, blockEl);
    for (const child of blockEl.childNodes) {
        buildRunsFromInline(child, fmt, runs);
    }
    if (runs.length === 0) runs.push(new TextRun({ text: '' }));
    return runs;
}

function getAlignment(blockEl) {
    const align = parseStyle(blockEl.getAttribute('style'))['text-align'];
    return align ? ALIGNMENT_MAP[align] : undefined;
}

function buildParagraphsFromList(listEl, baseFormat, reference, level, extraFirstOptions) {
    const paragraphs = [];
    let firstItemConsumed = false;
    for (const child of listEl.children) {
        if (child.tagName.toLowerCase() !== 'li') continue;

        const inlineNodes = [];
        const nestedLists = [];
        for (const liChild of child.childNodes) {
            if (liChild.nodeType === 1 && (liChild.tagName.toLowerCase() === 'ul' || liChild.tagName.toLowerCase() === 'ol')) {
                nestedLists.push(liChild);
            } else {
                inlineNodes.push(liChild);
            }
        }

        const runs = [];
        const fmt = inheritFormat(baseFormat, child);
        for (const node of inlineNodes) {
            buildRunsFromInline(node, fmt, runs);
        }
        if (runs.length === 0) runs.push(new TextRun({ text: '' }));

        const isFirst = !firstItemConsumed && extraFirstOptions;
        firstItemConsumed = true;
        paragraphs.push(new Paragraph({
            children: runs,
            numbering: { reference, level },
            ...(isFirst ? extraFirstOptions : {}),
        }));

        for (const nested of nestedLists) {
            const nestedRef = nested.tagName.toLowerCase() === 'ol' ? ORDERED_REF : BULLET_REF;
            paragraphs.push(...buildParagraphsFromList(nested, baseFormat, nestedRef, Math.min(level + 1, 1)));
        }
    }
    return paragraphs;
}

function buildParagraphsFromNode(node, baseFormat, extraFirstOptions) {
    if (node.nodeType === 3) {
        const text = node.textContent;
        if (!text || !text.trim()) return [];
        return [new Paragraph({
            children: [new TextRun({ text, ...toRunOptions(baseFormat) })],
            ...(extraFirstOptions || {}),
        })];
    }
    if (node.nodeType !== 1) return [];

    const tag = node.tagName.toLowerCase();

    if (tag === 'ul') return buildParagraphsFromList(node, baseFormat, BULLET_REF, 0, extraFirstOptions);
    if (tag === 'ol') return buildParagraphsFromList(node, baseFormat, ORDERED_REF, 0, extraFirstOptions);

    if (tag === 'br') {
        return [new Paragraph({ children: [new TextRun({ text: '' })], ...(extraFirstOptions || {}) })];
    }

    const runs = collectRuns(node, baseFormat);
    const alignment = getAlignment(node);
    return [new Paragraph({
        children: runs,
        ...(alignment ? { alignment } : {}),
        ...(extraFirstOptions || {}),
    })];
}

export function htmlToParagraphs(html, firstParagraphOptions) {
    if (!html || typeof html !== 'string') return [];
    const doc = new DOMParser().parseFromString(`<div id="root">${html}</div>`, 'text/html');
    const root = doc.getElementById('root');
    if (!root) return [];

    const paragraphs = [];
    let firstApplied = !firstParagraphOptions;
    for (const child of root.childNodes) {
        const opts = firstApplied ? undefined : firstParagraphOptions;
        const built = buildParagraphsFromNode(child, {}, opts);
        if (built.length > 0) {
            paragraphs.push(...built);
            firstApplied = true;
        }
    }
    return paragraphs;
}

export function buildEmailsDocument(emails) {
    const allParagraphs = [];
    (emails || []).forEach((email, idx) => {
        const firstOptions = idx > 0 ? { pageBreakBefore: true } : undefined;
        const paragraphs = htmlToParagraphs(email?.body || '', firstOptions);
        if (paragraphs.length === 0) {
            allParagraphs.push(new Paragraph({
                children: [new TextRun({ text: '' })],
                ...(firstOptions || {}),
            }));
        } else {
            allParagraphs.push(...paragraphs);
        }
    });
    if (allParagraphs.length === 0) {
        allParagraphs.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
    }

    return new Document({
        numbering: NUMBERING_CONFIG,
        sections: [{ children: allParagraphs }],
    });
}

export async function generateEmailsDocxBlob(emails) {
    const doc = buildEmailsDocument(emails || []);
    return Packer.toBlob(doc);
}

export async function downloadEmailsDocx(emails, filename = 'personalized-emails.docx') {
    const blob = await generateEmailsDocxBlob(emails);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
