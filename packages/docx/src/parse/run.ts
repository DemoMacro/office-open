import { findChild, findDeep, attr, attrNum, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import type { ParagraphChildJson } from "./types";

export function parseRun(run: Element, ctx: DocxParseContext): ParagraphChildJson | undefined {
    const rPr = findChild(run, "w:rPr");

    // Check for break
    const br = findChild(run, "w:br");
    if (br) {
        const brType = attr(br, "w:type");
        if (brType === "page") {
            return { $type: "pageBreak" };
        }
        // Other breaks (line, column) — treat as text with break marker
    }

    // Check for drawing (image)
    const drawing = findChild(run, "w:drawing") ?? findChild(run, "mc:AlternateContent");
    if (drawing) {
        const image = parseDrawingImage(drawing, ctx);
        if (image) return image;
    }

    // Check for pict (legacy image)
    const pict = findChild(run, "w:pict");
    if (pict) {
        const image = parsePictImage(pict, ctx);
        if (image) return image;
    }

    // Check for tab
    const tab = findChild(run, "w:tab");
    if (tab) {
        return { $type: "tab" };
    }

    // Check for footnote/endnote reference
    const footnoteRef = findChild(run, "w:footnoteReference");
    if (footnoteRef) {
        return undefined; // Skip for MVP
    }

    // Extract text
    const t = findChild(run, "w:t");
    const delText = findChild(run, "w:delText");
    const textEl = t ?? delText;
    const text = textOf(textEl);

    if (!text && !rPr) return undefined;

    const result: ParagraphChildJson = {
        $type: "textRun",
        text,
    };

    // Parse run properties
    if (rPr) {
        parseRunProperties(rPr, result);
    }

    return result;
}

function parseRunProperties(rPr: Element, out: Record<string, unknown>): void {
    // Bold
    const b = findChild(rPr, "w:b");
    if (b) {
        const val = attr(b, "w:val");
        out.bold = val !== "0" && val !== "false";
    }

    // Italic
    const i = findChild(rPr, "w:i");
    if (i) {
        const val = attr(i, "w:val");
        out.italics = val !== "0" && val !== "false";
    }

    // Underline
    const u = findChild(rPr, "w:u");
    if (u) {
        const val = attr(u, "w:val");
        const color = attr(u, "w:color");
        if (val && val !== "none" && val !== "false") {
            out.underline = { type: val };
            if (color) (out.underline as Record<string, unknown>).color = color;
        }
    }

    // Strike
    const strike = findChild(rPr, "w:strike");
    if (strike) {
        const val = attr(strike, "w:val");
        if (val !== "0" && val !== "false") out.strike = true;
    }

    // Double strike
    const dstrike = findChild(rPr, "w:dstrike");
    if (dstrike) {
        const val = attr(dstrike, "w:val");
        if (val !== "0" && val !== "false") out.doubleStrike = true;
    }

    // Small caps
    const smallCaps = findChild(rPr, "w:smallCaps");
    if (smallCaps) {
        const val = attr(smallCaps, "w:val");
        if (val !== "0" && val !== "false") out.smallCaps = true;
    }

    // All caps
    const caps = findChild(rPr, "w:caps");
    if (caps) {
        const val = attr(caps, "w:val");
        if (val !== "0" && val !== "false") out.allCaps = true;
    }

    // Subscript
    const sub = findChild(rPr, "w:vertAlign");
    if (sub) {
        const val = attr(sub, "w:val");
        if (val === "subscript") out.subScript = true;
        if (val === "superscript") out.superScript = true;
    }

    // Font size (half-points)
    const sz = findChild(rPr, "w:sz");
    if (sz) {
        const val = attrNum(sz, "w:val");
        if (val !== undefined) out.size = val;
    }

    // Color
    const color = findChild(rPr, "w:color");
    if (color) {
        const val = attr(color, "w:val");
        if (val && val !== "auto") out.color = val;
    }

    // Font family
    const rFonts = findChild(rPr, "w:rFonts");
    if (rFonts) {
        const ascii = attr(rFonts, "w:ascii");
        if (ascii) out.font = ascii;
    }

    // Highlight
    const highlight = findChild(rPr, "w:highlight");
    if (highlight) {
        const val = attr(highlight, "w:val");
        if (val && val !== "none") out.highlight = val;
    }

    // Character spacing (twips)
    const spacing = findChild(rPr, "w:spacing");
    if (spacing) {
        const val = attrNum(spacing, "w:val");
        if (val !== undefined) out.characterSpacing = val;
    }

    // Right-to-left
    const rtl = findChild(rPr, "w:rtl");
    if (rtl) {
        const val = attr(rtl, "w:val");
        if (val !== "0" && val !== "false") out.rightToLeft = true;
    }

    // Shading
    const shd = findChild(rPr, "w:shd");
    if (shd) {
        const fill = attr(shd, "w:fill");
        if (fill && fill !== "auto") {
            out.shading = { fill };
        }
    }

    // Kern
    const kern = findChild(rPr, "w:kern");
    if (kern) {
        const val = attrNum(kern, "w:val");
        if (val !== undefined) out.kern = val;
    }

    // Position
    const pos = findChild(rPr, "w:position");
    if (pos) {
        const val = attr(pos, "w:val");
        if (val !== undefined) out.position = val;
    }

    // Emboss
    const emboss = findChild(rPr, "w:emboss");
    if (emboss) {
        const val = attr(emboss, "w:val");
        if (val !== "0" && val !== "false") out.emboss = true;
    }

    // Imprint
    const imprint = findChild(rPr, "w:imprint");
    if (imprint) {
        const val = attr(imprint, "w:val");
        if (val !== "0" && val !== "false") out.imprint = true;
    }

    // Shadow
    const shadow = findChild(rPr, "w:shadow");
    if (shadow) {
        const val = attr(shadow, "w:val");
        if (val !== "0" && val !== "false") out.shadow = true;
    }

    // Outline
    const outline = findChild(rPr, "w:outline");
    if (outline) {
        const val = attr(outline, "w:val");
        if (val !== "0" && val !== "false") out.outline = true;
    }

    // Vanish
    const vanish = findChild(rPr, "w:vanish");
    if (vanish) {
        const val = attr(vanish, "w:val");
        if (val !== "0" && val !== "false") out.vanish = true;
    }

    // No proof
    const noProof = findChild(rPr, "w:noProof");
    if (noProof) {
        out.noProof = true;
    }

    // Math
    const math = findChild(rPr, "w:oMath");
    if (math) {
        out.math = true;
    }
}

function parseDrawingImage(
    drawing: Element,
    ctx: DocxParseContext,
): ParagraphChildJson | undefined {
    // Look for blip element (contains r:embed reference to image)
    const blips = findDeep(drawing, "a:blip");
    if (blips.length === 0) return undefined;

    const blip = blips[0];
    const embedId = attr(blip, "r:embed");
    if (!embedId) return undefined;

    // Find the media file from document relationships
    const rel = ctx.documentRels.get(embedId);
    if (!rel) return undefined;

    // Resolve media path
    const mediaPath = rel.target.startsWith("../")
        ? rel.target.replace("../", "word/")
        : `word/${rel.target}`;
    const mediaEntry = ctx.media.get(mediaPath);
    if (!mediaEntry) return undefined;

    // Extract dimensions
    const extent = findDeep(drawing, "wp:extent")[0] ?? findDeep(drawing, "wp:inline")[0];
    const cx = extent ? attrNum(extent, "cx") : undefined;
    const cy = extent ? attrNum(extent, "cy") : undefined;

    const result: ParagraphChildJson = {
        $type: "imageRun",
        type: mediaEntry.type,
        data: mediaEntry.data,
    };

    if (cx !== undefined || cy !== undefined) {
        result.transformation = {
            width: cx,
            height: cy,
        };
    }

    return result;
}

function parsePictImage(pict: Element, ctx: DocxParseContext): ParagraphChildJson | undefined {
    // Look for imagedata element with r:id
    const imageData = findDeep(pict, "v:imagedata")[0];
    if (!imageData) return undefined;

    const rid = attr(imageData, "r:id");
    if (!rid) return undefined;

    const rel = ctx.documentRels.get(rid);
    if (!rel) return undefined;

    const mediaPath = rel.target.startsWith("../")
        ? rel.target.replace("../", "word/")
        : `word/${rel.target}`;
    const mediaEntry = ctx.media.get(mediaPath);
    if (!mediaEntry) return undefined;

    return {
        $type: "imageRun",
        type: mediaEntry.type,
        data: mediaEntry.data,
    };
}
