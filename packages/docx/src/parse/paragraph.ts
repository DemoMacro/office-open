import { findChild, attr, attrNum, textOf, colorAttr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { parseRun } from "./run";
import { parseSdtRun } from "./sdt";
import type {
    ParagraphJson,
    ParagraphChildJson,
    ExternalHyperlinkJson,
    RawElement,
    TextRunJson,
    FieldJson,
} from "./types";
import { isRaw } from "./types";

export function parseParagraph(p: Element, ctx: DocxParseContext): ParagraphJson {
    const result: ParagraphJson = { $type: "paragraph" };
    const pPr = findChild(p, "w:pPr");

    if (pPr) {
        parseParagraphProperties(pPr, result);
    }

    // Extract children (runs, hyperlinks, etc.) with field code support
    const children: ParagraphChildJson[] = [];
    const elements = p.elements ?? [];
    let i = 0;

    while (i < elements.length) {
        const child = elements[i];

        if (child.name === "w:r") {
            // Check for field begin
            const fldChar = findChild(child, "w:fldChar");
            if (fldChar) {
                const fldCharType = attr(fldChar, "w:fldCharType");
                if (fldCharType === "begin") {
                    const field = collectField(elements, i, ctx);
                    if (field) children.push(field);
                    // Advance past the field
                    while (i < elements.length) {
                        const next = elements[i];
                        if (next.name === "w:r") {
                            const fc = findChild(next, "w:fldChar");
                            if (fc && attr(fc, "w:fldCharType") === "end") {
                                i++;
                                break;
                            }
                        }
                        i++;
                    }
                    continue;
                }
                // Skip separate/end fldChar runs (handled inside collectField)
                if (fldCharType === "separate" || fldCharType === "end") {
                    i++;
                    continue;
                }
            }
            const run = parseRun(child, ctx);
            if (run) children.push(run);
        } else if (child.name === "w:hyperlink") {
            const hyperlink = parseHyperlink(child, ctx);
            if (hyperlink) children.push(hyperlink);
        } else if (child.name === "w:bookmarkStart") {
            const name = attr(child, "w:name");
            if (name) children.push({ $type: "bookmark", name });
        } else if (child.name === "w:bookmarkEnd") {
            // Skip bookmark end markers
        } else if (child.name === "w:sdt") {
            const sdt = parseSdtRun(child, ctx);
            if (sdt) children.push(sdt);
            else children.push({ $raw: true, element: child } as RawElement);
        } else if (child.name === "w:oMath" || child.name === "w:oMathPara") {
            children.push({ $type: "math", element: child });
        } else if (child.name === "w:pPr") {
            // Paragraph properties — skip (already processed above)
        } else {
            children.push({ $raw: true, element: child } as RawElement);
        }
        i++;
    }

    // If only one text run with no formatting, simplify to text property
    const first = children[0];
    if (children.length === 1 && !isRaw(first) && first.$type === "textRun") {
        const textRun = first as TextRunJson;
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
        const val = attr(shd, "w:val");
        if (fill && fill !== "auto") {
            out.shading = { fill, ...(val && val !== "clear" && { type: val }) };
        }
    }

    // Paragraph borders (top, bottom, left, right, between)
    const pBdr = findChild(pPr, "w:pBdr");
    if (pBdr) {
        const borderSides = ["top", "bottom", "left", "right", "between"] as const;
        const borders: Record<string, unknown> = {};
        let hasNonNoneBorder = false;
        for (const side of borderSides) {
            const borderEl = findChild(pBdr, `w:${side}`);
            if (borderEl && borderEl.attributes && Object.keys(borderEl.attributes).length > 0) {
                const val = String(attr(borderEl, "w:val") ?? "");
                const sz = attrNum(borderEl, "w:sz");
                const color = colorAttr(borderEl, "w:color");
                const space = attrNum(borderEl, "w:space");
                if (val && val !== "none" && val !== "nil") {
                    hasNonNoneBorder = true;
                    const borderDef: Record<string, unknown> = { style: val };
                    if (sz !== undefined) borderDef.size = sz;
                    if (color) borderDef.color = color;
                    if (space !== undefined) borderDef.space = space;
                    borders[side] = borderDef;
                }
            }
        }
        if (Object.keys(borders).length > 0) out.border = borders;
        if (hasNonNoneBorder) out.thematicBreak = true;
    }

    // Tabs
    const tabs = findChild(pPr, "w:tabs");
    if (tabs) {
        const tabList: Array<{ pos: number; align?: string; leader?: string }> = [];
        for (const tab of tabs.elements ?? []) {
            if (tab.name === "w:tab") {
                const pos = attrNum(tab, "w:pos");
                const align = attr(tab, "w:val");
                const leader = attr(tab, "w:leader");
                if (pos !== undefined) {
                    tabList.push({
                        pos,
                        ...(align && align !== "left" && { align }),
                        ...(leader && leader !== "none" && { leader }),
                    });
                }
            }
        }
        if (tabList.length > 0) out.tabs = tabList;
    }

    // Suppress line numbers
    const suppressLineNumbers = findChild(pPr, "w:suppressLineNumbers");
    if (suppressLineNumbers) {
        const val = attr(suppressLineNumbers, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) out.suppressLineNumbers = true;
    }

    // Contextual spacing
    const contextualSpacing = findChild(pPr, "w:contextualSpacing");
    if (contextualSpacing) {
        const val = attr(contextualSpacing, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) out.contextualSpacing = true;
    }

    // Mirror indents
    const mirrorIndents = findChild(pPr, "w:mirrorIndents");
    if (mirrorIndents) {
        const val = attr(mirrorIndents, "w:val");
        if (val === undefined || (val !== "0" && val !== "false")) out.mirrorIndents = true;
    }
}

/** Collect field code (begin → separate → end) into a FieldJson */
function collectField(
    elements: Element[],
    startIndex: number,
    ctx: DocxParseContext,
): FieldJson | undefined {
    let instruction = "";
    let locked = false;
    let dirty = false;
    const fieldChildren: ParagraphChildJson[] = [];
    let pastSeparate = false;

    for (let i = startIndex + 1; i < elements.length; i++) {
        const el = elements[i];
        if (el.name !== "w:r") continue;

        const fldChar = findChild(el, "w:fldChar");
        if (fldChar) {
            const type = attr(fldChar, "w:fldCharType");
            if (type === "end") break;
            if (type === "separate") {
                pastSeparate = true;
                continue;
            }
        }

        if (fldChar) {
            // Check locked/dirty on begin char
            const fldLock = findChild(el, "w:fldChar");
            if (fldLock) {
                if (attr(fldLock, "w:fldLock") === "true") locked = true;
                if (attr(fldLock, "w:dirty") === "true") dirty = true;
            }
            continue;
        }

        if (!pastSeparate) {
            // Instruction part — collect instrText
            const instrText = findChild(el, "w:instrText");
            if (instrText) {
                instruction += textOf(instrText);
            }
        } else {
            // Result part — parse as normal run
            const run = parseRun(el, ctx);
            if (run) fieldChildren.push(run);
        }
    }

    const field: FieldJson = { $type: "field" };
    if (instruction) field.instruction = instruction;
    if (locked) field.locked = true;
    if (dirty) field.dirty = true;
    if (fieldChildren.length > 0) field.children = fieldChildren;

    return field;
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
