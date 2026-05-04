import { findChild, attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { parseRun } from "./run";
import type { ParagraphJson, ParagraphChildJson, ExternalHyperlinkJson } from "./types";

export function parseParagraph(p: Element, ctx: DocxParseContext): ParagraphJson {
    const result: ParagraphJson = { $type: "paragraph" };
    const pPr = findChild(p, "w:pPr");

    if (pPr) {
        parseParagraphProperties(pPr, result);
    }

    // Extract children (runs, hyperlinks, etc.)
    const children: ParagraphChildJson[] = [];

    for (const child of p.elements ?? []) {
        if (child.name === "w:r") {
            const run = parseRun(child, ctx);
            if (run) children.push(run);
        } else if (child.name === "w:hyperlink") {
            const hyperlink = parseHyperlink(child, ctx);
            if (hyperlink) children.push(hyperlink);
        } else if (child.name === "w:bookmarkStart") {
            const name = attr(child, "w:name");
            if (name) children.push({ $type: "bookmark", name });
        }
    }

    // If only one text run with no formatting, simplify to text property
    if (children.length === 1 && children[0].$type === "textRun") {
        const textRun = children[0] as Extract<ParagraphChildJson, { $type: "textRun" }>;
        const hasFormatting = Object.keys(textRun).some((k) => k !== "$type" && k !== "text");
        if (!hasFormatting && textRun.text !== undefined) {
            result.text = textRun.text;
        } else {
            result.children = children;
        }
    } else if (children.length > 0) {
        result.children = children;
    }

    return result;
}

function parseParagraphProperties(pPr: Element, out: ParagraphJson): void {
    // Style
    const pStyle = findChild(pPr, "w:pStyle");
    if (pStyle) {
        const styleVal = attr(pStyle, "w:val");
        if (styleVal) out.style = styleVal;
    }

    // Heading detection from style name
    if (out.style) {
        const headingMatch = out.style.match(/^HEADING_(\d+)$/);
        if (headingMatch) {
            out.heading = `HEADING_${headingMatch[1]}`;
        }
    }

    // Alignment
    const jc = findChild(pPr, "w:jc");
    if (jc) {
        const val = attr(jc, "w:val");
        if (val) out.alignment = val;
    }

    // Spacing
    const spacing = findChild(pPr, "w:spacing");
    if (spacing) {
        const spacingObj: ParagraphJson["spacing"] = {};
        const before = attrNum(spacing, "w:before");
        const after = attrNum(spacing, "w:after");
        const line = attrNum(spacing, "w:line");
        const lineRule = attr(spacing, "w:lineRule");

        if (before !== undefined) spacingObj.before = before;
        if (after !== undefined) spacingObj.after = after;
        if (line !== undefined) spacingObj.line = line;
        if (lineRule) spacingObj.lineRule = lineRule;

        if (Object.keys(spacingObj).length > 0) out.spacing = spacingObj;
    }

    // Indent
    const ind = findChild(pPr, "w:ind");
    if (ind) {
        const indentObj: ParagraphJson["indent"] = {};
        const left = attrNum(ind, "w:left") ?? attrNum(ind, "w:start");
        const right = attrNum(ind, "w:right") ?? attrNum(ind, "w:end");
        const firstLine = attrNum(ind, "w:firstLine");
        const hanging = attrNum(ind, "w:hanging");

        if (left !== undefined) indentObj.left = left;
        if (right !== undefined) indentObj.right = right;
        if (firstLine !== undefined) indentObj.firstLine = firstLine;
        if (hanging !== undefined) indentObj.hanging = hanging;

        if (Object.keys(indentObj).length > 0) out.indent = indentObj;
    }

    // Numbering
    const numPr = findChild(pPr, "w:numPr");
    if (numPr) {
        const numId = findChild(numPr, "w:numId");
        const ilvl = findChild(numPr, "w:ilvl");

        if (numId && ilvl) {
            const id = attr(numId, "w:val");
            const level = attrNum(ilvl, "w:val") ?? 0;
            if (id && id !== "0") {
                out.numbering = { reference: id, level };
            }
        }
    }

    // Outline level
    const outlineLvl = findChild(pPr, "w:outlineLvl");
    if (outlineLvl) {
        const val = attrNum(outlineLvl, "w:val");
        if (val !== undefined) out.outlineLevel = val;
    }

    // Keep next
    const keepNext = findChild(pPr, "w:keepNext");
    if (keepNext) {
        const val = attr(keepNext, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) out.keepNext = true;
    }

    // Keep lines
    const keepLines = findChild(pPr, "w:keepLines");
    if (keepLines) {
        const val = attr(keepLines, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) out.keepLines = true;
    }

    // Widow control
    const widowControl = findChild(pPr, "w:widowControl");
    if (widowControl) {
        const val = attr(widowControl, "w:val");
        if (val === "0" || val === "false") out.widowControl = false;
    }

    // Page break before
    const pageBreakBefore = findChild(pPr, "w:pageBreakBefore");
    if (pageBreakBefore) {
        const val = attr(pageBreakBefore, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) out.pageBreakBefore = true;
    }

    // Shading
    const shd = findChild(pPr, "w:shd");
    if (shd) {
        const fill = attr(shd, "w:fill");
        if (fill && fill !== "auto") {
            out.shading = { fill };
        }
    }

    // Thematic break (bottom border on paragraph)
    const pBdr = findChild(pPr, "w:pBdr");
    if (pBdr) {
        const bottom = findChild(pBdr, "w:bottom");
        if (bottom) {
            const val = attr(bottom, "w:val");
            if (val && val !== "none" && val !== "nil") {
                out.thematicBreak = true;
            }
        }
    }
}

function parseHyperlink(hl: Element, ctx: DocxParseContext): ExternalHyperlinkJson | undefined {
    const rid = attr(hl, "r:id");
    let link = "";

    if (rid) {
        link = ctx.hyperlinks.get(rid) ?? "";
    }

    // Also check anchor (internal link)
    const anchor = attr(hl, "w:anchor");
    if (!link && anchor) {
        link = `#${anchor}`;
    }

    if (!link) return undefined;

    const tooltip = attr(hl, "w:tooltip");

    // Parse hyperlink children (runs)
    const children: ParagraphChildJson[] = [];
    for (const child of hl.elements ?? []) {
        if (child.name === "w:r") {
            const run = parseRun(child, ctx);
            if (run) children.push(run);
        }
    }

    const result: ExternalHyperlinkJson = {
        $type: "externalHyperlink",
        link,
        ...(tooltip && { tooltip }),
        ...(children.length > 0 && { children }),
    };

    return result;
}
