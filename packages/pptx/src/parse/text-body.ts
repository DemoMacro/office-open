import { findChild, children, attr, attrBool, attrNum, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PptxParseContext } from "./context";
import type { ParagraphJson, RunJson } from "./types";

export function parseTextBody(
    txBody: Element,
    ctx: PptxParseContext,
): {
    paragraphs?: ParagraphJson[];
    vertical?: string;
    anchor?: string;
    autoFit?: string;
    wrap?: string;
    margins?: Record<string, number>;
    columns?: number;
    columnSpacing?: number;
} {
    const result: ReturnType<typeof parseTextBody> = {};

    // Body properties
    const bodyPr = findChild(txBody, "a:bodyPr");
    if (bodyPr) {
        const vert = attr(bodyPr, "vert");
        if (vert) result.vertical = vert;

        const anchor = attr(bodyPr, "anchor");
        if (anchor) result.anchor = anchor;

        const wrap = attr(bodyPr, "wrap");
        if (wrap) result.wrap = wrap;

        const lIns = attrNum(bodyPr, "lIns");
        const tIns = attrNum(bodyPr, "tIns");
        const rIns = attrNum(bodyPr, "rIns");
        const bIns = attrNum(bodyPr, "bIns");
        if (lIns !== undefined || tIns !== undefined || rIns !== undefined || bIns !== undefined) {
            const margins: Record<string, number> = {};
            if (lIns !== undefined) margins.left = lIns;
            if (tIns !== undefined) margins.top = tIns;
            if (rIns !== undefined) margins.right = rIns;
            if (bIns !== undefined) margins.bottom = bIns;
            result.margins = margins;
        }

        const numCol = attrNum(bodyPr, "numCol");
        if (numCol !== undefined && numCol > 1) result.columns = numCol;

        const spcCol = attrNum(bodyPr, "spcCol");
        if (spcCol !== undefined) result.columnSpacing = spcCol;

        // AutoFit
        const normAutofit = findChild(bodyPr, "a:normAutofit");
        const spAutoFit = findChild(bodyPr, "a:spAutoFit");
        const noAutofit = findChild(bodyPr, "a:noAutofit");

        if (normAutofit) result.autoFit = "normal";
        if (spAutoFit) result.autoFit = "shape";
        if (noAutofit) result.autoFit = "none";
    }

    // Parse paragraphs
    const paragraphs = children(txBody, "a:p");
    if (paragraphs.length > 0) {
        result.paragraphs = paragraphs.map((p) => parsePptxParagraph(p, ctx));
    }

    return result;
}

export function parsePptxParagraph(p: Element, ctx: PptxParseContext): ParagraphJson {
    const result: ParagraphJson = {};
    const pPr = findChild(p, "a:pPr");

    if (pPr) {
        const algn = attr(pPr, "algn");
        if (algn) result.alignment = algn;

        const lvl = attrNum(pPr, "lvl");
        const marL = attrNum(pPr, "marL");
        const indent = attrNum(pPr, "indent");
        if (lvl !== undefined || marL !== undefined || indent !== undefined) {
            const indentObj: Record<string, number> = {};
            if (lvl !== undefined) indentObj.level = lvl;
            if (marL !== undefined) indentObj.left = marL;
            if (indent !== undefined) indentObj.margin = indent;
            result.indent = indentObj;
        }

        const spcBef = findChild(pPr, "a:spcBef");
        if (spcBef) {
            const spcPts = findChild(spcBef, "a:spcPts");
            if (spcPts) {
                const val = attrNum(spcPts, "val");
                if (val !== undefined) {
                    if (!result.spacing) result.spacing = {};
                    result.spacing.before = val;
                }
            }
        }

        const spcAft = findChild(pPr, "a:spcAft");
        if (spcAft) {
            const spcPts = findChild(spcAft, "a:spcPts");
            if (spcPts) {
                const val = attrNum(spcPts, "val");
                if (val !== undefined) {
                    if (!result.spacing) result.spacing = {};
                    result.spacing.after = val;
                }
            }
        }

        const lnSpc = findChild(pPr, "a:lnSpc");
        if (lnSpc) {
            const spcPts = findChild(lnSpc, "a:spcPts");
            if (spcPts) {
                const val = attrNum(spcPts, "val");
                if (val !== undefined) {
                    if (!result.spacing) result.spacing = {};
                    result.spacing.line = val;
                }
            }
        }

        // Bullet/numbering detection
        const buNone = findChild(pPr, "a:buNone");
        if (!buNone) {
            const buChar = findChild(pPr, "a:buChar");
            const buAutoNum = findChild(pPr, "a:buAutoNum");

            if (buAutoNum) {
                const type = attr(buAutoNum, "type");
                result.numbering = { type: type ?? "arabicPeriod", level: lvl ?? 0 };
            } else if (buChar) {
                const char = attr(buChar, "char");
                result.bullet = { type: char ?? "bullet", level: lvl ?? 0 };
            } else if (lvl !== undefined && lvl > 0) {
                // Has indent level but no explicit bullet — likely inherits from master
                result.bullet = { level: lvl };
            }
        }
    }

    // Parse runs
    const runs: RunJson[] = [];

    for (const child of p.elements ?? []) {
        if (child.name === "a:r") {
            const run = parsePptxRun(child, ctx);
            if (run) runs.push(run);
        } else if (child.name === "a:br") {
            runs.push({ text: "\n" });
        } else if (child.name === "a:fld") {
            // Field element — extract text from t element
            const t = findChild(child, "a:t");
            if (t) {
                const text = textOf(t);
                if (text) runs.push({ text });
            }
        } else if (child.name === "a:hlinkClick") {
            // Hyperlink run
            const r = children(child, "a:r")[0];
            if (r) {
                const run = parsePptxRun(r, ctx);
                if (run) {
                    const action = attr(child, "action");
                    if (action) {
                        run.hyperlink = { url: action };
                    }
                    const tooltip = attr(child, "tooltip");
                    if (tooltip && run.hyperlink) {
                        run.hyperlink.tooltip = tooltip;
                    }
                    runs.push(run);
                }
            }
        }
    }

    if (runs.length > 0) {
        result.children = runs;
    }

    // If only one text run with no formatting, simplify
    if (runs.length === 1 && runs[0].text !== undefined) {
        const hasFormatting = Object.keys(runs[0]).some((k) => k !== "text");
        if (!hasFormatting) {
            result.text = runs[0].text;
            delete result.children;
        }
    }

    return result;
}

function parsePptxRun(r: Element, _ctx: PptxParseContext): RunJson | undefined {
    const result: RunJson = {};
    const rPr = findChild(r, "a:rPr");

    // Extract text
    const t = findChild(r, "a:t");
    const text = textOf(t);
    if (text) result.text = text;

    if (!text && !rPr) return undefined;

    if (rPr) {
        const sz = attrNum(rPr, "sz");
        if (sz !== undefined) result.fontSize = sz;

        const b = attrBool(rPr, "b");
        if (b) result.bold = true;

        const i = attrBool(rPr, "i");
        if (i) result.italic = true;

        const u = attr(rPr, "u");
        if (u && u !== "none") result.underline = u;

        const strike = attr(rPr, "strike");
        if (strike && strike !== "noStrike") result.strike = strike;

        const lang = attr(rPr, "lang");
        if (lang) result.lang = lang;

        const spc = attrNum(rPr, "spc");
        if (spc !== undefined) result.spacing = spc;

        const baseline = attrNum(rPr, "baseline");
        if (baseline !== undefined) result.baseline = baseline;

        const cap = attr(rPr, "cap");
        if (cap) result.capitalization = cap;

        const dirty = attr(rPr, "dirty");
        if (dirty === "1") result.noProof = true; // dirty flag often used with noProof

        const noProof = attr(rPr, "noProof");
        if (noProof === "1") result.noProof = true;

        const rtl = attr(rPr, "rtl");
        if (rtl === "1") result.rightToLeft = true;

        // Font family
        const latin = findChild(rPr, "a:latin");
        if (latin) {
            const typeface = attr(latin, "typeface");
            if (typeface) result.font = typeface;
        }

        // Fill (solid color)
        const solidFill = findChild(rPr, "a:solidFill");
        if (solidFill) {
            const srgbClr = findChild(solidFill, "a:srgbClr");
            const schemeClr = findChild(solidFill, "a:schemeClr");
            if (srgbClr) {
                const val = attr(srgbClr, "val");
                if (val) result.fill = { type: "solid", color: val };
            } else if (schemeClr) {
                const val = attr(schemeClr, "val");
                if (val) result.fill = { type: "solid", color: val };
            }
        }
    }

    return result;
}
