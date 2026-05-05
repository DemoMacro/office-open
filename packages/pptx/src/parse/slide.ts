import { readXmlFromZip } from "@office-open/core";
import { findChild, attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PptxParseContext } from "./context";
import { getSlideRels } from "./context";
import { parseShape, parsePicture, parseConnector, parsePptxTable } from "./shape";
import { parseTextBody } from "./text-body";
import type { SlideJson, SlideChildJson, GroupShapeJson, ParagraphJson } from "./types";

export function parseSlide(slideXml: Element, ctx: PptxParseContext, slidePath: string): SlideJson {
    const result: SlideJson = { children: [] };

    const slideRels = getSlideRels(ctx, slidePath);

    const cSld = findChild(slideXml, "p:cSld");
    if (!cSld) return result;

    // Background
    const bg = findChild(cSld, "p:bg");
    if (bg) {
        result.background = parseSlideBackground(bg);
    }

    // Shape tree
    const spTree = findChild(cSld, "p:spTree");
    if (spTree) {
        for (const child of spTree.elements ?? []) {
            // Skip spTree's own group properties (not child shapes)
            if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
            const parsed = parseShapeTreeElement(child, ctx, slideRels);
            if (parsed) result.children.push(parsed);
        }
    }

    // Transition (enhanced)
    const transition = findChild(slideXml, "p:transition");
    if (transition) {
        result.transition = parseTransition(transition);
    }

    // Notes
    let notesTarget: string | undefined;
    for (const [, rel] of slideRels) {
        if (rel.type.includes("notesSlide")) {
            notesTarget = rel.target;
            break;
        }
    }
    if (notesTarget) {
        const notesPath = notesTarget.startsWith("../")
            ? notesTarget.replace("../", "ppt/")
            : `ppt/${notesTarget}`;
        const notesXml = readXmlFromZip(ctx.zip, notesPath);
        if (notesXml) {
            const notesContent = parseSlideNotes(notesXml, ctx);
            if (notesContent) result.notes = notesContent;
        }
    }

    return result;
}

function parseShapeTreeElement(
    el: Element,
    ctx: PptxParseContext,
    slideRels: Map<string, { target: string; type: string }>,
): SlideChildJson | undefined {
    switch (el.name) {
        case "p:sp":
            return parseShape(el, ctx, slideRels);
        case "p:pic":
            return parsePicture(el, ctx, slideRels);
        case "p:cxnSp":
            return parseConnector(el, ctx);
        case "p:graphicFrame": {
            const table = parsePptxTable(el, ctx);
            return table ?? { $raw: true, element: el };
        }
        case "p:grpSp":
            return parseGroupShape(el, ctx, slideRels);
        default:
            return { $raw: true, element: el };
    }
}

/** Recursively parse group shapes */
function parseGroupShape(
    grpSp: Element,
    ctx: PptxParseContext,
    slideRels: Map<string, { target: string; type: string }>,
): GroupShapeJson {
    const result: GroupShapeJson = { $type: "groupShape" };

    const nvSpPr = findChild(grpSp, "p:nvGrpSpPr");
    const cNvPr = findChild(nvSpPr, "p:cNvPr") ?? findChild(nvSpPr, "p:nvPr");
    if (cNvPr) {
        const id = attrNum(cNvPr, "id");
        const name = attr(cNvPr, "name");
        if (id !== undefined) result.id = id;
        if (name) result.name = name;
    }

    const grpSpPr = findChild(grpSp, "p:grpSpPr");
    if (grpSpPr) {
        const xfrm = findChild(grpSpPr, "a:xfrm");
        if (xfrm) {
            const off = findChild(xfrm, "a:off");
            const ext = findChild(xfrm, "a:ext");
            if (off) {
                result.x = attrNum(off, "x");
                result.y = attrNum(off, "y");
            }
            if (ext) {
                result.width = attrNum(ext, "cx");
                result.height = attrNum(ext, "cy");
            }
            const rot = attrNum(xfrm, "rot");
            if (rot !== undefined && rot !== 0) {
                // OOXML rotation is in 60000ths of a degree
                result.rotation = rot / 60000;
            }
            const flipH = attr(xfrm, "flipH");
            const flipV = attr(xfrm, "flipV");
            if (flipH === "1") result.flipH = true;
            if (flipV === "1") result.flipV = true;
        }
    }

    // Parse children — CT_GroupShape allows sp/grpSp/etc directly (no spTree wrapper)
    // Some producers use spTree inside grpSp; handle both layouts
    const childSource = findChild(grpSp, "p:spTree") ?? grpSp;
    const children: SlideChildJson[] = [];
    for (const child of childSource.elements ?? []) {
        if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
        const parsed = parseShapeTreeElement(child, ctx, slideRels);
        if (parsed) children.push(parsed);
    }
    if (children.length > 0) result.children = children;

    return result;
}

/** Parse slide notes content */
function parseSlideNotes(
    notesXml: Element,
    ctx: PptxParseContext,
): string | ParagraphJson[] | undefined {
    const cSld = findChild(notesXml, "p:cSld");
    if (!cSld) return undefined;

    const spTree = findChild(cSld, "p:spTree");
    if (!spTree) return undefined;

    for (const child of spTree.elements ?? []) {
        if (child.name === "p:sp") {
            const nvSpPr = findChild(child, "p:nvSpPr");
            const nvPr = findChild(nvSpPr, "p:nvPr");
            if (nvPr) {
                const ph = findChild(nvPr, "p:ph");
                if (ph && (attr(ph, "type") === "body" || attr(ph, "type") === undefined)) {
                    const txBody = findChild(child, "p:txBody");
                    if (txBody) {
                        const parsed = parseTextBody(txBody, ctx);
                        if (parsed.paragraphs && parsed.paragraphs.length > 0) {
                            return parsed.paragraphs;
                        }
                    }
                }
            }
        }
    }
    return undefined;
}

function parseSlideBackground(bg: Element): string | Record<string, unknown> | undefined {
    const bgPr = findChild(bg, "p:bgPr");
    if (!bgPr) return undefined;

    // Check for solid fill
    const solidFill = findChild(bgPr, "a:solidFill");
    if (solidFill) {
        const srgbClr = findChild(solidFill, "a:srgbClr");
        const schemeClr = findChild(solidFill, "a:schemeClr");
        const color = srgbClr
            ? attr(srgbClr, "val")
            : schemeClr
              ? attr(schemeClr, "val")
              : undefined;
        if (color) return color;
    }

    // Gradient fill
    const gradFill = findChild(bgPr, "a:gradFill");
    if (gradFill) {
        return { type: "gradient" };
    }

    return undefined;
}

function parseTransition(transition: Element): Record<string, unknown> {
    const result: Record<string, unknown> = {};

    const spd = attr(transition, "spd");
    const advClick = attr(transition, "advClick");
    const advTm = attrNum(transition, "advTm");
    const p14Dur = attrNum(transition, "p14:dur");

    if (spd) result.speed = spd;
    if (advClick !== undefined) result.advanceOnClick = advClick === "1";
    if (advTm !== undefined) result.advanceAfterMs = advTm;
    if (p14Dur !== undefined) result.durationMs = p14Dur;

    // Get the first child element which defines the transition type
    for (const child of transition.elements ?? []) {
        if (child.name?.startsWith("p:") && child.name !== "p:transition") {
            const type = child.name.replace("p:", "");
            result.type = type;

            // Transition-specific options
            const p14Prst = findChild(child, "p14:prst");
            if (p14Prst) {
                const prstVal = attr(p14Prst, "val");
                if (prstVal) result.preset = prstVal;
            }
            break;
        }
    }

    return result;
}
