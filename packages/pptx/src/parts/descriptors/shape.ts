/**
 * Shape (p:sp) and Picture (p:pic) descriptors for PPTX slides.
 *
 * These CustomDescriptor implementations produce the same XML output as the
 * class-based Shape and Picture components, but through the descriptor pipeline.
 *
 * @module
 */

import { convertEmuToPixels, convertToEmu, parseOnOff, toUint8Array } from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions, ShapeLockingOptions } from "@office-open/core";
import type { CustomDescriptor, WriteContext, ReadContext } from "@office-open/core/descriptor";
import { parse } from "@office-open/core/descriptor";
import {
  shapeLockingDesc,
  effectListDesc,
  shapePropertiesDesc,
  textBodyDesc,
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
  createSolidFill,
  stringifyColorChoice,
} from "@office-open/core/drawing";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild, findFirst, escapeXml, attrNum, attr } from "@office-open/xml";
import { imageTypeFromPath } from "@shared/media/image-type";
import type { PictureOptions } from "@shared/picture";
import { readShapeStyle, type ShapeOptions, type ShapeStyleOptions } from "@shared/shape/shape";

import type { PptxWriteContext, MediaEntry } from "../../context";

// ── Auto-incrementing IDs ──

let _nextShapeId = 2;
let _nextPictureId = 100;

/** Reset shape ID counter (useful for tests). */
export function resetShapeIdCounter(value = 2): void {
  _nextShapeId = value;
}

/** Reset picture ID counter (useful for tests). */
export function resetPictureIdCounter(value = 100): void {
  _nextPictureId = value;
}

// ── Shape (p:sp) descriptor ──

export const shapeDesc: CustomDescriptor<ShapeOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextShapeId++;
    const name = opts.name ?? `Shape ${id}`;
    const parts: string[] = [];

    // ── p:nvSpPr ──
    parts.push(stringifyNvSpPr(id, name, opts));

    // ── p:spPr ──
    const spPrXml = stringifySpPr(opts, ctx);
    if (spPrXml) parts.push(spPrXml);

    // ── p:style ──
    if (opts.style) {
      const styleXml = stringifyShapeStyle(opts.style, ctx);
      if (styleXml) parts.push(styleXml);
    }

    // ── p:txBody (optional in CT_Shape; omitted when absent so txBody-less
    // shapes like the notes sldImg placeholder round-trip without a spurious
    // empty body. Parse only sets textBody when a p:txBody element exists.)
    if (opts.textBody !== undefined) {
      parts.push(`<p:txBody>${textBodyDesc.stringify(opts.textBody, ctx) ?? ""}</p:txBody>`);
    }

    // ── Root attributes ──
    const spAttrs: string[] = [];
    if (opts.useBackgroundFill) spAttrs.push(' useBgFill="1"');
    if (opts.blackWhiteMode) spAttrs.push(` bwMode="${opts.blackWhiteMode}"`);

    return `<p:sp${spAttrs.join("")}>${parts.join("")}</p:sp>`;
  },

  parse(el, ctx) {
    const result: ShapeOptions = {};

    // Root attributes
    if (el.attributes) {
      if (el.attributes["useBgFill"] !== undefined)
        result.useBackgroundFill = parseOnOff(el.attributes["useBgFill"]) ?? false;
      if (el.attributes["bwMode"] !== undefined)
        result.blackWhiteMode = String(el.attributes["bwMode"]) as ShapeOptions["blackWhiteMode"];
    }

    // p:nvSpPr
    const nvSpPr = findChild(el, "p:nvSpPr");
    if (nvSpPr) {
      const parsed = readNvSpPr(nvSpPr);
      if (parsed.id !== undefined) result.id = parsed.id;
      if (parsed.name !== undefined) result.name = parsed.name;
      if (parsed.placeholder !== undefined) result.placeholder = parsed.placeholder;
      if (parsed.placeholderIndex !== undefined) result.placeholderIndex = parsed.placeholderIndex;
      if (parsed.placeholderSize !== undefined) result.placeholderSize = parsed.placeholderSize;
      if (parsed.placeholderOrientation !== undefined)
        result.placeholderOrientation = parsed.placeholderOrientation;
      if (parsed.hasCustomPrompt !== undefined) result.hasCustomPrompt = parsed.hasCustomPrompt;
      if (parsed.isPhoto !== undefined) result.isPhoto = parsed.isPhoto;
      if (parsed.userDrawn !== undefined) result.userDrawn = parsed.userDrawn;
      if (parsed.locking !== undefined) result.locking = parsed.locking;
    }

    // p:spPr
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      Object.assign(result, readSpPr(spPr, ctx));
    }

    // p:style
    const style = findChild(el, "p:style");
    if (style) result.style = readShapeStyle(style);

    // p:txBody
    const txBody = findChild(el, "p:txBody");
    if (txBody) result.textBody = textBodyDesc.parse(txBody, ctx);

    return result;
  },
};

// ── Picture (p:pic) descriptor ──

export const pictureDesc: CustomDescriptor<PictureOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextPictureId++;
    const name = opts.name ?? `Picture ${id}`;
    const fileName = `${name.replace(/\s+/g, "_")}.${opts.type}`;

    // Geometry: number is already EMU, string is UniversalMeasure → EMU
    const widthEmu = convertToEmu(opts.width ?? 0);
    const heightEmu = convertToEmu(opts.height ?? 0);

    // Register media with the PPTX context (content-deduplicated)
    const mediaEntry = pptx.addImage(fileName, {
      key: fileName,
      type: opts.type,
      fileName,
      data: toUint8Array(opts.data, { encoding: "base64" }),
      transformation: {
        pixels: {
          x: Math.round(convertEmuToPixels(widthEmu)),
          y: Math.round(convertEmuToPixels(heightEmu)),
        },
        emus: { x: widthEmu, y: heightEmu },
      },
    });

    const parts: string[] = [];

    // ── p:nvPicPr ──
    parts.push(stringifyNvPicPr(id, name, opts));

    // ── p:blipFill ──
    parts.push(stringifyPptxBlipFill(mediaEntry.fileName));

    // ── p:spPr ──
    const spPrXml = stringifyPicSpPr(opts, ctx);
    if (spPrXml) parts.push(spPrXml);

    return `<p:pic>${parts.join("")}</p:pic>`;
  },

  parse(el, ctx) {
    const result: Partial<PictureOptions> = {};

    // p:nvPicPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "p:cNvPr");
      Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
      if (cNvPr?.attributes && cNvPr.attributes["id"] !== undefined) {
        result.id = Number(cNvPr.attributes["id"]);
      }
    }

    // p:spPr (position/size only)
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) {
        const off = findChild(xfrm, "a:off");
        if (off?.attributes) {
          result.x = Number(off.attributes["x"] ?? 0);
          result.y = Number(off.attributes["y"] ?? 0);
        }
        const ext = findChild(xfrm, "a:ext");
        if (ext?.attributes) {
          result.width = Number(ext.attributes["cx"] ?? 0);
          result.height = Number(ext.attributes["cy"] ?? 0);
        }
      }

      // Effects (e.g. shadow/reflection on the picture)
      const effectLst = findChild(spPr, "a:effectLst");
      if (effectLst) {
        const effects = parse(effectListDesc, effectLst, ctx);
        if (effects) result.effects = effects;
      }
    }

    // Image data from p:blipFill → a:blip → r:embed
    const blip = findFirst(el, "a:blip");
    if (blip) {
      const rEmbed = attr(blip, "r:embed");
      if (rEmbed) {
        const imagePath = ctx.resolveRelationship(rEmbed);
        if (imagePath) {
          const imageData = ctx.getRaw(imagePath);
          if (imageData) {
            result.data = imageData;
            result.type = imageTypeFromPath(imagePath);
          }
        }
      }
    }

    // Defaults if image data could not be resolved
    if (!result.data) result.data = new Uint8Array(0);
    if (!result.type) result.type = "png";

    return result as PictureOptions;
  },
};

// ── Shape helper: p:nvSpPr ──

function stringifyNvSpPr(id: number, name: string, opts: ShapeOptions): string {
  // nvPr
  let nvPrContent = "<p:nvPr/>";
  if (opts.placeholder) {
    const phAttrs: string[] = [`type="${opts.placeholder}"`];
    if (opts.placeholderIndex !== undefined) phAttrs.push(`idx="${opts.placeholderIndex}"`);
    if (opts.placeholderSize !== undefined) phAttrs.push(`sz="${opts.placeholderSize}"`);
    if (opts.placeholderOrientation !== undefined)
      phAttrs.push(`orient="${opts.placeholderOrientation}"`);
    if (opts.hasCustomPrompt) phAttrs.push('hasCustomPrompt="1"');
    nvPrContent = `<p:nvPr><p:ph ${phAttrs.join(" ")}/></p:nvPr>`;
  } else if (opts.isPhoto || opts.userDrawn) {
    const nvPrAttrs: string[] = [];
    if (opts.isPhoto) nvPrAttrs.push('isPhoto="1"');
    if (opts.userDrawn) nvPrAttrs.push('userDrawn="1"');
    nvPrContent = `<p:nvPr ${nvPrAttrs.join(" ")}/>`;
  }

  // cNvSpPr (with optional locking)
  let cNvSpPrContent = "<p:cNvSpPr/>";
  if (opts.locking) {
    const lockAttrs = buildLockAttrs(opts.locking);
    if (lockAttrs.length > 0) {
      cNvSpPrContent = `<p:cNvSpPr><a:spLocks ${lockAttrs.join(" ")}/></p:cNvSpPr>`;
    }
  }

  const cNvPrXml = stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name);
  return `<p:nvSpPr>${cNvPrXml}${cNvSpPrContent}${nvPrContent}</p:nvSpPr>`;
}

function buildLockAttrs(opts: ShapeLockingOptions): string[] {
  const attrs: string[] = [];
  const keys = [
    "noGrp",
    "noSelect",
    "noRot",
    "noChangeAspect",
    "noMove",
    "noResize",
    "noEditPoints",
    "noAdjustHandles",
    "noChangeArrowheads",
    "noChangeShapeType",
    "noTextEdit",
  ] as const;
  for (const key of keys) {
    const val = opts[key];
    if (val !== undefined) attrs.push(`${key}="${val ? 1 : 0}"`);
  }
  return attrs;
}

// ── Shape helper: p:spPr ──

function stringifySpPr(opts: ShapeOptions, ctx: WriteContext): string {
  const pptx = ctx as PptxWriteContext;

  // Blip fill: pre-register via addImage so the deduped canonical keeps the
  // pptx-local fileName (image_blip.<type>). shapePropertiesDesc's fillDesc
  // re-registers via addMedia, deduped to the same canonical.
  if (opts.fill && typeof opts.fill !== "string" && opts.fill.type === "blip" && opts.fill.data) {
    const blipFill = opts.fill;
    const raw = toUint8Array(blipFill.data, { encoding: "base64" });
    const fileName = `image_blip.${blipFill.imageType ?? "png"}`;
    pptx.addImage(fileName, {
      key: fileName,
      data: raw,
      fileName,
      type: (blipFill.imageType ?? "png") as MediaEntry["type"],
      transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
    });
  }

  // Placeholder shapes inherit geometry/fill from layout/master; only emit
  // them when explicitly set. Non-placeholder shapes default to rect geometry
  // and noFill.
  const isPlaceholder = !!opts.placeholder;
  const geometry = opts.customGeometry
    ? undefined
    : isPlaceholder
      ? opts.geometry
      : (opts.geometry ?? "rect");
  const fill = isPlaceholder ? opts.fill : (opts.fill ?? ({ type: "none" } as const));

  const spPrContent = shapePropertiesDesc.stringify(
    {
      x: opts.x,
      y: opts.y,
      width: opts.width,
      height: opts.height,
      flipHorizontal: opts.flipHorizontal,
      rotation: opts.rotation,
      customGeometry: opts.customGeometry,
      geometry,
      fill,
      outline: opts.outline,
      effects: opts.effects,
      scene3d: opts.scene3d,
      shape3d: opts.shape3d,
      ext: opts.ext,
    },
    ctx,
  );

  if (!spPrContent) return "<p:spPr/>";
  return `<p:spPr>${spPrContent}</p:spPr>`;
}

// ── Shape helper: p:style ──

/**
 * Serialize p:style (CT_ShapeStyle). lnRef/fillRef/effectRef carry a bare
 * EG_ColorChoice; fontRef wraps it in a:solidFill. Shared by the shape
 * descriptor and the master placeholder emitter.
 */
export function stringifyShapeStyle(style: ShapeStyleOptions, ctx: WriteContext): string {
  const ref = (
    name: string,
    { index, color }: { index: number; color?: string },
    wrapSolidFill: boolean,
  ): string => {
    if (color === undefined) return `<${name} idx="${index}"/>`;
    const colorChoice = wrapSolidFill
      ? createSolidFill({ value: color })
      : stringifyColorChoice({ value: color }, ctx);
    return `<${name} idx="${index}">${colorChoice}</${name}>`;
  };

  const parts: string[] = [];
  if (style.lineReference) parts.push(ref("a:lnRef", style.lineReference, false));
  if (style.fillReference) parts.push(ref("a:fillRef", style.fillReference, false));
  if (style.effectReference) parts.push(ref("a:effectRef", style.effectReference, false));
  if (style.fontReference) parts.push(ref("a:fontRef", style.fontReference, true));
  return parts.length > 0 ? `<p:style>${parts.join("")}</p:style>` : "";
}

// ── Picture helpers ──

function stringifyNvPicPr(id: number, name: string, opts?: PictureOptions): string {
  const cNvPrXml = stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name);
  return `<p:nvPicPr>${cNvPrXml}<p:cNvPicPr/><p:nvPr/></p:nvPicPr>`;
}

/** PPTX uses p:blipFill (not pic:blipFill). */
function stringifyPptxBlipFill(fileName: string): string {
  return `<p:blipFill><a:blip r:embed="{${escapeXml(fileName)}}" cstate="none"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
}

function stringifyPicSpPr(opts: PictureOptions, ctx: WriteContext): string {
  const spPrContent = shapePropertiesDesc.stringify(
    {
      x: opts.x,
      y: opts.y,
      width: opts.width,
      height: opts.height,
      // Pictures always use a rect preset geometry.
      geometry: "rect",
      effects: opts.effects,
    },
    ctx,
  );
  if (!spPrContent) return "<p:spPr/>";
  return `<p:spPr>${spPrContent}</p:spPr>`;
}

// ── Read helpers ──

/**
 * Read the non-visual drawing properties (id/name/description/title/hidden…)
 * from the first parent tag that holds a p:cNvPr (a:cNvPr tolerated); with no
 * parent tags, el itself is the non-visual properties container. Shared
 * probe+parse by every slide-child descriptor.
 */
export function readCnvPr(
  el: XmlElement,
  ...parentTags: string[]
): NonVisualDrawingPropertiesOptions & { id?: number } {
  const parents = parentTags.length ? parentTags.map((tag) => findChild(el, tag)) : [el];
  for (const parent of parents) {
    if (!parent) continue;
    const cNvPr = findChild(parent, "p:cNvPr") ?? findChild(parent, "a:cNvPr");
    if (!cNvPr) continue;
    const result: NonVisualDrawingPropertiesOptions & { id?: number } = {
      ...parseNonVisualDrawingProperties(cNvPr),
    };
    const id = attrNum(cNvPr, "id");
    if (id !== undefined) result.id = id;
    return result;
  }
  return {};
}

export function readNvSpPr(nvSpPr: XmlElement): ShapeOptions {
  const result: ShapeOptions = {};

  Object.assign(result, readCnvPr(nvSpPr));

  const nvPr = findChild(nvSpPr, "p:nvPr");
  if (nvPr) {
    if (nvPr.attributes) {
      if (nvPr.attributes["isPhoto"] !== undefined)
        result.isPhoto = parseOnOff(nvPr.attributes["isPhoto"]) ?? false;
      if (nvPr.attributes["userDrawn"] !== undefined)
        result.userDrawn = parseOnOff(nvPr.attributes["userDrawn"]) ?? false;
    }
    const ph = findChild(nvPr, "p:ph");
    if (ph?.attributes) {
      if (ph.attributes["type"] !== undefined)
        result.placeholder = String(ph.attributes["type"]) as ShapeOptions["placeholder"];
      if (ph.attributes["idx"] !== undefined)
        result.placeholderIndex = Number(ph.attributes["idx"]);
      if (ph.attributes["sz"] !== undefined)
        result.placeholderSize = ph.attributes["sz"] as ShapeOptions["placeholderSize"];
      if (ph.attributes["orient"] !== undefined)
        result.placeholderOrientation = ph.attributes[
          "orient"
        ] as ShapeOptions["placeholderOrientation"];
      if (ph.attributes["hasCustomPrompt"] !== undefined)
        result.hasCustomPrompt = parseOnOff(ph.attributes["hasCustomPrompt"]) ?? false;
    }
  }

  const cNvSpPr = findChild(nvSpPr, "p:cNvSpPr");
  if (cNvSpPr) {
    const spLocks = findChild(cNvSpPr, "a:spLocks");
    if (spLocks) {
      result.locking = shapeLockingDesc.parse(spLocks, {} as ReadContext) as ShapeLockingOptions;
    }
  }

  return result;
}

/** Parse p:spPr via the shared core descriptor. */
function readSpPr(spPr: XmlElement, ctx: ReadContext): ShapeOptions {
  const result = parse(shapePropertiesDesc, spPr, ctx) as ShapeOptions;

  // Collapse a preset geometry without adjustment values to the bare string
  // shorthand the public options accept.
  const geometry = result.geometry;
  if (
    typeof geometry === "object" &&
    (!geometry.adjustmentValues || geometry.adjustmentValues.length === 0) &&
    geometry.preset
  ) {
    result.geometry = geometry.preset;
  }
  return result;
}

/** Read x/y/width/height (in EMU) from an a:xfrm element. */
export function readPositionFromXfrm(xfrm: XmlElement): Record<string, number> {
  const result: Record<string, number> = {};
  const off = findChild(xfrm, "a:off");
  if (off) {
    const x = attrNum(off, "x");
    if (x !== undefined) result.x = x;
    const y = attrNum(off, "y");
    if (y !== undefined) result.y = y;
  }
  const ext = findChild(xfrm, "a:ext");
  if (ext) {
    const cx = attrNum(ext, "cx");
    if (cx !== undefined) result.width = cx;
    const cy = attrNum(ext, "cy");
    if (cy !== undefined) result.height = cy;
  }
  return result;
}

// ── Image type helper ──
