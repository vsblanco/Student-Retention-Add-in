import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    ExternalHyperlink,
    AlignmentType,
    SimpleMailMergeField,
    ImageRun,
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
const FIELD_RE = /\{(\w+)\}/g;
const ALIGNMENT_MAP = {
    left: AlignmentType.LEFT,
    right: AlignmentType.RIGHT,
    center: AlignmentType.CENTER,
    justify: AlignmentType.JUSTIFIED,
};
const SUPPORTED_IMAGE_TYPES = ['png', 'jpg', 'gif', 'bmp'];

export function extractFieldNames(html) {
    if (!html || typeof html !== 'string') return [];
    const seen = new Set();
    const names = [];
    for (const match of html.matchAll(FIELD_RE)) {
        if (!seen.has(match[1])) {
            seen.add(match[1]);
            names.push(match[1]);
        }
    }
    return names;
}

function normalizeImageType(rawType) {
    if (!rawType) return null;
    const t = rawType.toLowerCase();
    if (t === 'jpeg') return 'jpg';
    return SUPPORTED_IMAGE_TYPES.includes(t) ? t : null;
}

function parseDataUrl(src) {
    const match = src.match(/^data:image\/([\w+-]+)(?:;charset=[^;,]+)?;base64,(.+)$/i);
    if (!match) return null;
    const type = normalizeImageType(match[1]);
    if (!type) return null;
    try {
        const bin = atob(match[2]);
        const data = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
        return { type, data };
    } catch {
        return null;
    }
}

function blobToArrayBuffer(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
    });
}

async function fetchRemoteImage(src) {
    try {
        const resp = await fetch(src);
        if (!resp.ok) return null;
        const blob = await resp.blob();
        const buffer = await blobToArrayBuffer(blob);

        let type = normalizeImageType((blob.type.match(/^image\/([\w+-]+)/) || [])[1]);
        if (!type) {
            const ext = (src.split('?')[0].split('#')[0].split('.').pop() || '').toLowerCase();
            type = normalizeImageType(ext);
        }
        if (!type) return null;
        return { type, data: new Uint8Array(buffer) };
    } catch (err) {
        console.warn(`[docxGenerator] failed to fetch image ${src}:`, err);
        return null;
    }
}

function measureImage(src, fallback) {
    return new Promise(resolve => {
        if (typeof Image === 'undefined') {
            resolve(fallback);
            return;
        }
        const img = new Image();
        const done = (w, h) => resolve({ width: w || fallback.width, height: h || fallback.height });
        img.onload = () => done(img.naturalWidth, img.naturalHeight);
        img.onerror = () => done(0, 0);
        img.src = src;
    });
}

async function prepareImageMap(html) {
    if (!html || typeof html !== 'string' || !html.includes('<img')) return new Map();
    const doc = new DOMParser().parseFromString(`<div>${html}</div>`, 'text/html');
    const imgs = Array.from(doc.querySelectorAll('img'));
    const map = new Map();

    await Promise.all(imgs.map(async img => {
        const src = img.getAttribute('src');
        if (!src || map.has(src)) return;

        const imgData = src.startsWith('data:')
            ? parseDataUrl(src)
            : await fetchRemoteImage(src);
        if (!imgData) return;

        const attrW = parseInt(img.getAttribute('width'), 10);
        const attrH = parseInt(img.getAttribute('height'), 10);
        let width = Number.isFinite(attrW) && attrW > 0 ? attrW : 0;
        let height = Number.isFinite(attrH) && attrH > 0 ? attrH : 0;

        if (!width || !height) {
            const fallback = { width: width || 300, height: height || 150 };
            const measured = await measureImage(src, fallback);
            width = width || measured.width;
            height = height || measured.height;
        }

        map.set(src, { ...imgData, width, height });
    }));

    return map;
}

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

function pushTextSegment(text, format, runs) {
    if (!text) return;
    const runOpts = toRunOptions(format);
    FIELD_RE.lastIndex = 0;
    let lastIdx = 0;
    let match;
    while ((match = FIELD_RE.exec(text)) !== null) {
        if (match.index > lastIdx) {
            runs.push(new TextRun({ text: text.slice(lastIdx, match.index), ...runOpts }));
        }
        runs.push(new SimpleMailMergeField(match[1]));
        lastIdx = FIELD_RE.lastIndex;
    }
    if (lastIdx < text.length) {
        runs.push(new TextRun({ text: text.slice(lastIdx), ...runOpts }));
    }
}

function buildImageRun(node, imageMap) {
    const src = node.getAttribute('src');
    const data = src && imageMap?.get(src);
    if (!data) return null;
    return new ImageRun({
        data: data.data,
        type: data.type,
        transformation: { width: data.width, height: data.height },
    });
}

function buildRunsFromInline(node, format, runs, imageMap) {
    if (node.nodeType === 3) {
        pushTextSegment(node.textContent, format, runs);
        return;
    }
    if (node.nodeType !== 1) return;

    const tag = node.tagName.toLowerCase();

    if (tag === 'br') {
        runs.push(new TextRun({ break: 1, ...toRunOptions(format) }));
        return;
    }

    if (tag === 'img') {
        const run = buildImageRun(node, imageMap);
        if (run) runs.push(run);
        return;
    }

    if (tag === 'a') {
        const href = node.getAttribute('href');
        const linkFormat = inheritFormat({ ...format, underline: true, color: format.color || '0563C1' }, node);
        const linkRuns = [];
        for (const child of node.childNodes) {
            buildRunsFromInline(child, linkFormat, linkRuns, imageMap);
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
        buildRunsFromInline(child, childFormat, runs, imageMap);
    }
}

function collectRuns(blockEl, baseFormat, imageMap) {
    const runs = [];
    const fmt = inheritFormat(baseFormat, blockEl);
    for (const child of blockEl.childNodes) {
        buildRunsFromInline(child, fmt, runs, imageMap);
    }
    if (runs.length === 0) runs.push(new TextRun({ text: '' }));
    return runs;
}

function getAlignment(blockEl) {
    const align = parseStyle(blockEl.getAttribute('style'))['text-align'];
    return align ? ALIGNMENT_MAP[align] : undefined;
}

function buildParagraphsFromList(listEl, baseFormat, reference, level, imageMap) {
    const paragraphs = [];
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
            buildRunsFromInline(node, fmt, runs, imageMap);
        }
        if (runs.length === 0) runs.push(new TextRun({ text: '' }));
        paragraphs.push(new Paragraph({ children: runs, numbering: { reference, level } }));

        for (const nested of nestedLists) {
            const nestedRef = nested.tagName.toLowerCase() === 'ol' ? ORDERED_REF : BULLET_REF;
            paragraphs.push(...buildParagraphsFromList(nested, baseFormat, nestedRef, Math.min(level + 1, 1), imageMap));
        }
    }
    return paragraphs;
}

function buildParagraphsFromNode(node, baseFormat, options) {

    if (node.nodeType === 3) {
        const text = node.textContent;
        if (!text || !text.trim()) return [];
        const runs = [];
        pushTextSegment(text, baseFormat, runs);
        return [new Paragraph({ children: runs })];
    }
    if (node.nodeType !== 1) return [];

    const tag = node.tagName.toLowerCase();

    if (tag === 'ul') return buildParagraphsFromList(node, baseFormat, BULLET_REF, 0, options.imageMap);
    if (tag === 'ol') return buildParagraphsFromList(node, baseFormat, ORDERED_REF, 0, options.imageMap);

    if (tag === 'br') {
        return [new Paragraph({ children: [new TextRun({ text: '' })] })];
    }

    const runs = collectRuns(node, baseFormat, options.imageMap);
    const alignment = getAlignment(node);
    return [new Paragraph({
        children: runs,
        ...(alignment ? { alignment } : {}),
    })];
}

export function htmlToParagraphs(html, options = {}) {
    if (!html || typeof html !== 'string') return [];
    const doc = new DOMParser().parseFromString(`<div id="root">${html}</div>`, 'text/html');
    const root = doc.getElementById('root');
    if (!root) return [];

    const opts = {
        imageMap: options.imageMap || new Map()
    };

    const paragraphs = [];
    for (const child of root.childNodes) {
        paragraphs.push(...buildParagraphsFromNode(child, {}, opts));
    }
    return paragraphs;
}

export function buildMailMergeTemplate(bodyHtml, options = {}) {
    const paragraphs = htmlToParagraphs(bodyHtml || '', options);
    if (paragraphs.length === 0) {
        paragraphs.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
    }
    return new Document({
        numbering: NUMBERING_CONFIG,
        sections: [{ children: paragraphs }],
    });
}

export async function generateMailMergeTemplateBlob(bodyHtml, options = {}) {
    const imageMap = options.imageMap || await prepareImageMap(bodyHtml);
    const doc = buildMailMergeTemplate(bodyHtml, {
        imageMap
    });
    return Packer.toBlob(doc);
}

export async function downloadMailMergeTemplate(bodyHtml, filename = 'email-template.docx', options = {}) {
    const blob = await generateMailMergeTemplateBlob(bodyHtml, options);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
