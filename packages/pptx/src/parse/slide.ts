import { findChild, attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PptxParseContext } from "./context";
import { getSlideRels } from "./context";
import { parseShape, parsePicture, parseConnector, parsePptxTable } from "./shape";
import type { SlideJson, SlideChildJson } from "./types";

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
            const parsed = parseShapeTreeElement(child, ctx, slideRels);
            if (parsed) result.children.push(parsed);
        }
    }

    // Transition
    const transition = findChild(slideXml, "p:transition");
    if (transition) {
        result.transition = parseTransition(transition);
    }

    // Notes
    // Notes are in separate files, skipped for MVP

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
        case "p:graphicFrame":
            return parsePptxTable(el, ctx);
        case "p:grpSp":
            return undefined; // Group shapes — skip for MVP
        default:
            return undefined;
    }
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

    // Get the first child element which defines the transition type
    for (const child of transition.elements ?? []) {
        if (child.name?.startsWith("p:")) {
            const type = child.name.replace("p:", "");
            const spd = attr(transition, "spd");
            const advClick = attr(transition, "advClick");

            result.type = type;
            if (spd) result.speed = spd;
            if (advClick !== undefined) result.advanceOnClick = advClick === "1";
            break;
        }
    }

    return result;
}
