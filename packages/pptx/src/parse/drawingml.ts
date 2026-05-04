import { findChild, attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

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

export function parseFill(spPr: Element): Record<string, unknown> | undefined {
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

        const angle = attrNum(gradFill, "lin")
            ? attrNum(findChild(gradFill, "a:lin")!, "ang")
            : undefined;

        return {
            type: "gradient",
            stops,
            ...(angle !== undefined && { angle }),
        };
    }

    const noFill = findChild(spPr, "a:noFill");
    if (noFill) return { type: "none" };

    // Pattern fill, blip fill — simplified for MVP
    const pattFill = findChild(spPr, "a:pattFill");
    if (pattFill) return { type: "pattern" };

    const blipFill = findChild(spPr, "a:blipFill");
    if (blipFill) return { type: "blip" };

    return undefined;
}

export function parseOutline(spPr: Element): Record<string, unknown> | undefined {
    const ln = findChild(spPr, "a:ln");
    if (!ln) return undefined;

    const result: Record<string, unknown> = {};

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

    // No line
    const noFill = findChild(ln, "a:noFill");
    if (noFill) result.width = 0;

    return Object.keys(result).length > 0 ? result : undefined;
}

export function parseEffects(spPr: Element): Record<string, unknown> | undefined {
    const effectLst = findChild(spPr, "a:effectLst");
    if (!effectLst) return undefined;

    const result: Record<string, unknown> = {};

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
        result.innerShadow = {};
    }

    // Glow
    const glow = findChild(effectLst, "a:glow");
    if (glow) {
        const radius = attrNum(glow, "rad");
        if (radius !== undefined) result.glow = { radius };
    }

    // Reflection
    const reflection = findChild(effectLst, "a:reflection");
    if (reflection) {
        result.reflection = {};
    }

    return Object.keys(result).length > 0 ? result : undefined;
}
