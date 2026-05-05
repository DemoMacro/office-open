import { findChild, findDeep, attr, attrNum, textOf, colorAttr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { getMediaData } from "./context";
import type { ParagraphChildJson } from "./types";

export function parseRun(run: Element, ctx: DocxParseContext): ParagraphChildJson | undefined {
    const rPr = findChild(run, "w:rPr");

    // Check for break
    const br = findChild(run, "w:br");
    if (br) {
        const brType = attr(br, "w:type");
        if (brType === "page" || brType === undefined) {
            return { $type: "pageBreak" };
        }
        if (brType === "column") {
            return { $type: "columnBreak" };
        }
        if (brType === "line" || brType === "textWrapping") {
            return { $type: "lineBreak" };
        }
        // section break handled at paragraph level
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
    const endnoteRef = findChild(run, "w:endnoteReference");
    if (footnoteRef || endnoteRef) {
        // Preserve as raw for round-trip
        const refEl = footnoteRef ?? endnoteRef!;
        return { $raw: true, element: refEl } as ParagraphChildJson;
    }

    // Extract text
    const t = findChild(run, "w:t");
    const delText = findChild(run, "w:delText");
    const textEl = t ?? delText;
    const text = textOf(textEl);

    if (!text && !rPr) return undefined;

    const result: any = {
        $type: "textRun",
        text,
        ...(delText && !t && { deletedText: true }),
    };

    // Parse run properties
    if (rPr) {
        parseRunProperties(rPr, result);
    }

    return result as ParagraphChildJson;
}

export function parseRunProperties(rPr: Element, out: Record<string, unknown>): void {
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

    // Bold complex script
    const bCs = findChild(rPr, "w:bCs");
    if (bCs) {
        const val = attr(bCs, "w:val");
        if (val !== "0" && val !== "false") out.boldCs = true;
    }

    // Italic complex script
    const iCs = findChild(rPr, "w:iCs");
    if (iCs) {
        const val = attr(iCs, "w:val");
        if (val !== "0" && val !== "false") out.italicCs = true;
    }

    // Underline
    const u = findChild(rPr, "w:u");
    if (u) {
        const val = attr(u, "w:val");
        const color = colorAttr(u, "w:color");
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

    // Complex script font size
    const szCs = findChild(rPr, "w:szCs");
    if (szCs) {
        const val = attrNum(szCs, "w:val");
        if (val !== undefined) out.sizeCs = val;
    }

    // Color
    const colorEl = findChild(rPr, "w:color");
    if (colorEl) {
        const val = colorAttr(colorEl, "w:val");
        if (val && val !== "auto") out.color = val;
    }

    // Font family (complete)
    const rFonts = findChild(rPr, "w:rFonts");
    if (rFonts) {
        const ascii = attr(rFonts, "w:ascii");
        const hAnsi = attr(rFonts, "w:hAnsi");
        const eastAsia = attr(rFonts, "w:eastAsia");
        const cs = attr(rFonts, "w:cs");
        const hint = attr(rFonts, "w:hint");

        if (hAnsi || eastAsia || cs || hint) {
            out.font = {
                ...(ascii && { ascii }),
                ...(hAnsi && { hAnsi }),
                ...(eastAsia && { eastAsia }),
                ...(cs && { cs }),
                ...(hint && hint !== "default" && { hint }),
            };
        } else if (ascii) {
            out.font = ascii;
        }
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
        const val = attr(shd, "w:val");
        if (fill && fill !== "auto") {
            out.shading = { fill, ...(val && val !== "clear" && { type: val }) };
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

    // Language
    const lang = findChild(rPr, "w:lang");
    if (lang) {
        const val = attr(lang, "w:val");
        if (val) out.lang = val;
    }

    // Math
    const math = findChild(rPr, "w:oMath");
    if (math) {
        out.math = true;
    }
}

export function parseDrawingImage(
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
    const mediaEntry = getMediaData(ctx, mediaPath);
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

export function parsePictImage(
    pict: Element,
    ctx: DocxParseContext,
): ParagraphChildJson | undefined {
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
    const mediaEntry = getMediaData(ctx, mediaPath);
    if (!mediaEntry) return undefined;

    return {
        $type: "imageRun",
        type: mediaEntry.type,
        data: mediaEntry.data,
    };
}
