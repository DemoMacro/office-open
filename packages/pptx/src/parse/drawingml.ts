import type { FillOptions } from "@office-open/core/drawingml";
import { findChild, attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParsedOutlineOptions, ParsedEffectsOptions } from "./types";

export function parseTransform(spPr: Element): {
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
} {
    const result: ReturnType<typeof parseTransform> = {};

    const xfrm = findChild(spPr, "a:xfrm");
    if (xfrm) {
        const off = findChild(xfrm, "a:off");
        if (off) {
            result.x = attrNum(off, "x");
            result.y = attrNum(off, "y");
        }

        const ext = findChild(xfrm, "a:ext");
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
        if (flipH === "1") result.flipH = true;

        const flipV = attr(xfrm, "flipV");
        if (flipV === "1") result.flipV = true;
    }

    return result;
}

export function parseGeometry(spPr: Element): string | undefined {
    const prstGeom = findChild(spPr, "a:prstGeom");
    if (prstGeom) {
        return attr(prstGeom, "prst") ?? undefined;
    }
    return undefined;
}

export function parseFill(spPr: Element): FillOptions | undefined {
    const solidFill = findChild(spPr, "a:solidFill");
    if (solidFill) {
        const srgbClr = findChild(solidFill, "a:srgbClr");
        const schemeClr = findChild(solidFill, "a:schemeClr");

        let color: string | undefined;
        if (srgbClr) {
            color = attr(srgbClr, "val");
        } else if (schemeClr) {
            color = attr(schemeClr, "val");
        }

        if (color) return { type: "solid", color };
    }

    const gradFill = findChild(spPr, "a:gradFill");
    if (gradFill) {
        const stops: Array<{ position: number; color: string }> = [];
        const gsLst = findChild(gradFill, "a:gsLst");
        if (gsLst) {
            for (const gs of gsLst.elements ?? []) {
                if (gs.name !== "a:gs") continue;
                const pos = attrNum(gs, "pos") ?? 0;
                const srgbClr = findChild(gs, "a:srgbClr");
                const schemeClr = findChild(gs, "a:schemeClr");
                const color = srgbClr
                    ? attr(srgbClr, "val")
                    : schemeClr
                      ? attr(schemeClr, "val")
                      : undefined;
                if (color) stops.push({ position: pos, color });
            }
        }

        const result: Record<string, unknown> = { type: "gradient", stops };

        // Linear gradient
        const lin = findChild(gradFill, "a:lin");
        if (lin) {
            const angle = attrNum(lin, "ang");
            const scaled = attr(lin, "scaled");
            if (angle !== undefined) result.angle = angle;
            if (scaled !== undefined) result.scaled = scaled === "1";
        }

        // Path gradient
        const path = findChild(gradFill, "a:path");
        if (path) {
            result.path = attr(path, "path") ?? "circle";
            const fillToRect = findChild(path, "a:fillToRect");
            if (fillToRect) {
                result.fillToRect = {
                    l: attr(fillToRect, "l"),
                    t: attr(fillToRect, "t"),
                    r: attr(fillToRect, "r"),
                    b: attr(fillToRect, "b"),
                };
            }
        }

        return result as FillOptions;
    }

    const noFill = findChild(spPr, "a:noFill");
    if (noFill) return { type: "none" };

    // Pattern fill — complete
    const pattFill = findChild(spPr, "a:pattFill");
    if (pattFill) {
        const prst = attr(pattFill, "prst");
        const fgClr = findChild(pattFill, "a:fgClr");
        const bgClr = findChild(pattFill, "a:bgClr");
        const fg = fgClr
            ? findChild(fgClr, "a:srgbClr")
                ? attr(findChild(fgClr, "a:srgbClr")!, "val")
                : findChild(fgClr, "a:schemeClr")
                  ? attr(findChild(fgClr, "a:schemeClr")!, "val")
                  : undefined
            : undefined;
        const bg = bgClr
            ? findChild(bgClr, "a:srgbClr")
                ? attr(findChild(bgClr, "a:srgbClr")!, "val")
                : findChild(bgClr, "a:schemeClr")
                  ? attr(findChild(bgClr, "a:schemeClr")!, "val")
                  : undefined
            : undefined;
        return { type: "pattern", pattern: prst!, ...(fg && { fg }), ...(bg && { bg }) };
    }

    // Blip fill — complete
    const blipFill = findChild(spPr, "a:blipFill");
    if (blipFill) {
        const result: Record<string, unknown> = { type: "blip" };
        const blip = findChild(blipFill, "a:blip");
        if (blip) {
            const embed = attr(blip, "r:embed");
            const link = attr(blip, "r:link");
            if (embed) result.embed = embed;
            if (link) result.link = link;
        }
        // Stretch
        const stretch = findChild(blipFill, "a:stretch");
        if (stretch) {
            result.stretch = true;
            const fillRect = findChild(stretch, "a:fillRect");
            if (fillRect) {
                result.fillRect = {
                    l: attr(fillRect, "l"),
                    t: attr(fillRect, "t"),
                    r: attr(fillRect, "r"),
                    b: attr(fillRect, "b"),
                };
            }
        }
        // Tile
        const tile = findChild(blipFill, "a:tile");
        if (tile) {
            delete result.stretch;
            result.tile = true;
        }
        return result as FillOptions;
    }

    return undefined;
}

export function parseOutline(spPr: Element): ParsedOutlineOptions | undefined {
    const ln = findChild(spPr, "a:ln");
    if (!ln) return undefined;

    const result: ParsedOutlineOptions = {};

    const w = attrNum(ln, "w");
    if (w !== undefined) result.width = w;

    const cap = attr(ln, "cap");
    if (cap) result.cap = cap;

    const cmpd = attr(ln, "cmpd");
    if (cmpd) result.compound = cmpd;

    const algn = attr(ln, "algn");
    if (algn) result.alignment = algn;

    // Fill
    const solidFill = findChild(ln, "a:solidFill");
    if (solidFill) {
        const srgbClr = findChild(solidFill, "a:srgbClr");
        const schemeClr = findChild(solidFill, "a:schemeClr");
        const color = srgbClr
            ? attr(srgbClr, "val")
            : schemeClr
              ? attr(schemeClr, "val")
              : undefined;
        if (color) result.color = color;
    }

    // Dash style
    const prstDash = findChild(ln, "a:prstDash");
    if (prstDash) {
        const val = attr(prstDash, "val");
        if (val) result.dashStyle = val;
    }

    // Round/bevel
    const round = findChild(ln, "a:round");
    if (round) result.round = true;

    const bevel = findChild(ln, "a:bevel");
    if (bevel) result.bevel = true;

    // Line head/tail end
    const headEnd = findChild(ln, "a:headEnd");
    if (headEnd) {
        const he: Record<string, unknown> = {};
        const type = attr(headEnd, "type");
        const w = attrNum(headEnd, "w");
        const len = attrNum(headEnd, "len");
        if (type) he.type = type;
        if (w !== undefined) he.width = w;
        if (len !== undefined) he.length = len;
        if (Object.keys(he).length > 0) result.headEnd = he;
    }

    const tailEnd = findChild(ln, "a:tailEnd");
    if (tailEnd) {
        const te: Record<string, unknown> = {};
        const type = attr(tailEnd, "type");
        const w = attrNum(tailEnd, "w");
        const len = attrNum(tailEnd, "len");
        if (type) te.type = type;
        if (w !== undefined) te.width = w;
        if (len !== undefined) te.length = len;
        if (Object.keys(te).length > 0) result.tailEnd = te;
    }

    // Join style
    const join = findChild(ln, "a:join");
    if (join) {
        const type = attr(join, "type");
        if (type) result.join = type;
    }

    // No line
    const noFill = findChild(ln, "a:noFill");
    if (noFill) result.width = 0;

    return Object.keys(result).length > 0 ? result : undefined;
}

export function parseEffects(spPr: Element): ParsedEffectsOptions | undefined {
    const effectLst = findChild(spPr, "a:effectLst");
    if (!effectLst) return undefined;

    const result: ParsedEffectsOptions = {};

    // Outer shadow
    const outerShdw = findChild(effectLst, "a:outerShdw");
    if (outerShdw) {
        const shadow: Record<string, unknown> = {};
        const blurRad = attrNum(outerShdw, "blurRad");
        const dist = attrNum(outerShdw, "dist");
        const dir = attrNum(outerShdw, "dir");
        const algn = attr(outerShdw, "algn");
        const rotWithShape = attr(outerShdw, "rotWithShape");

        if (blurRad !== undefined) shadow.blurRadius = blurRad;
        if (dist !== undefined) shadow.distance = dist;
        if (dir !== undefined) shadow.direction = dir;
        if (algn) shadow.alignment = algn;
        if (rotWithShape !== undefined) shadow.rotateWithShape = rotWithShape === "1";

        const solidFill = findChild(outerShdw, "a:solidFill");
        if (solidFill) {
            const srgbClr = findChild(solidFill, "a:srgbClr");
            if (srgbClr) shadow.color = attr(srgbClr, "val");
        }

        result.outerShadow = shadow;
    }

    // Inner shadow
    const innerShdw = findChild(effectLst, "a:innerShdw");
    if (innerShdw) {
        const shadow: Record<string, unknown> = {};
        const blurRad = attrNum(innerShdw, "blurRad");
        const dist = attrNum(innerShdw, "dist");
        const dir = attrNum(innerShdw, "dir");
        const rotWithShape = attr(innerShdw, "rotWithShape");
        if (blurRad !== undefined) shadow.blurRadius = blurRad;
        if (dist !== undefined) shadow.distance = dist;
        if (dir !== undefined) shadow.direction = dir;
        if (rotWithShape !== undefined) shadow.rotateWithShape = rotWithShape === "1";
        const solidFill = findChild(innerShdw, "a:solidFill");
        if (solidFill) {
            const srgbClr = findChild(solidFill, "a:srgbClr");
            if (srgbClr) shadow.color = attr(srgbClr, "val");
        }
        result.innerShadow = shadow;
    }

    // Glow
    const glow = findChild(effectLst, "a:glow");
    if (glow) {
        const radius = attrNum(glow, "rad");
        const glowObj: Record<string, unknown> = {};
        if (radius !== undefined) glowObj.radius = radius;
        const solidFill = findChild(glow, "a:solidFill");
        if (solidFill) {
            const srgbClr = findChild(solidFill, "a:srgbClr");
            if (srgbClr) glowObj.color = attr(srgbClr, "val");
        }
        if (Object.keys(glowObj).length > 0) result.glow = glowObj;
    }

    // Reflection
    const reflection = findChild(effectLst, "a:reflection");
    if (reflection) {
        const refObj: Record<string, unknown> = {};
        const blurRad = attrNum(reflection, "blurRad");
        const dist = attrNum(reflection, "dist");
        const dir = attrNum(reflection, "dir");
        const fadeDir = attrNum(reflection, "fadeDir");
        const stA = attrNum(reflection, "stA");
        const stB = attrNum(reflection, "stB");
        const rotWithShape = attr(reflection, "rotWithShape");
        if (blurRad !== undefined) refObj.blurRadius = blurRad;
        if (dist !== undefined) refObj.distance = dist;
        if (dir !== undefined) refObj.direction = dir;
        if (fadeDir !== undefined) refObj.fadeDirection = fadeDir;
        if (stA !== undefined) refObj.startAlpha = stA;
        if (stB !== undefined) refObj.endAlpha = stB;
        if (rotWithShape !== undefined) refObj.rotateWithShape = rotWithShape === "1";
        if (Object.keys(refObj).length > 0) result.reflection = refObj;
    }

    return Object.keys(result).length > 0 ? result : undefined;
}
