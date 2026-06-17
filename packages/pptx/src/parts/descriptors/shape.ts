/**
 * Shape (p:sp) and Picture (p:pic) descriptors for PPTX slides.
 *
 * These CustomDescriptor implementations produce the same XML output as the
 * class-based Shape and Picture components, but through the descriptor pipeline.
 *
 * @module
 */

import {
  convertPixelsToEmu,
  convertToEmu,
  xsdRectAlignment,
  xsdTextAnchor,
} from "@office-open/core";
import type { ShapeLockingOptions } from "@office-open/core";
import type { CustomDescriptor, WriteContext, ReadContext } from "@office-open/core/descriptor";
import { stringify, parse } from "@office-open/core/descriptor";
import {
  transform2DDesc,
  fillDesc,
  presetGeometryDesc,
  outlineDesc,
  shapeLockingDesc,
  effectListDesc,
  scene3DDesc,
  shape3DDesc,
} from "@office-open/core/drawingml";
import type {
  Transform2DOptions,
  FillOptions as CoreFillOptions,
  OutlineOptions as CoreOutlineOptions,
} from "@office-open/core/drawingml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild, findDeep, escapeXml, attrNum, attr } from "@office-open/xml";
import type { EffectsOptions, ReflectionOptions } from "@shared/drawingml/effects";
import type { OutlineOptions } from "@shared/drawingml/outline";
import type { ParagraphOptions } from "@shared/shape/paragraph/paragraph";

import type { PptxWriteContext, MediaEntry } from "../../context";
import { paragraphDesc, type ParagraphDescriptorOptions } from "./text";

// ── Types ──

export interface ShapeStyleDescriptorOptions {
  lineReference?: { index: number; color?: string };
  fillReference?: { index: number; color?: string };
  effectReference?: { index: number; color?: string };
  fontReference?: { index: number; color?: string };
}

export interface TextBodyDescriptorOptions {
  text?: string;
  children?: (ParagraphOptions | string)[];
  vertical?:
    | "horz"
    | "vert"
    | "vert270"
    | "wordArtVert"
    | "eaVert"
    | "mongolianVert"
    | "wordArtVertRtl";
  anchor?: "top" | "center" | "bottom" | "justify" | "distribute";
  autoFit?: "normal" | "shape" | "none";
  wrap?: "square" | "none";
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  marginTop?: number;
  marginBottom?: number;
  columns?: number;
  columnSpacing?: number;
}

export interface ShapeDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  geometry?: string;
  fill?: CoreFillOptions;
  outline?: OutlineOptions;
  effects?: EffectsOptions;
  flipHorizontal?: boolean;
  rotation?: number;
  textBody?: TextBodyDescriptorOptions;
  locking?: ShapeLockingOptions;
  placeholder?: "title" | "body" | "subTitle" | "sldNum" | "dt" | "ftr" | "hdr" | "obj";
  placeholderIndex?: number;
  useBackgroundFill?: boolean;
  isPhoto?: boolean;
  userDrawn?: boolean;
  hasCustomPrompt?: boolean;
  style?: ShapeStyleDescriptorOptions;
  blackWhiteMode?:
    | "clr"
    | "auto"
    | "gray"
    | "ltGray"
    | "invGray"
    | "grayWhite"
    | "blackGray"
    | "blackWhite"
    | "black"
    | "white"
    | "hidden";
}

export interface PictureDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  data: Uint8Array;
  type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
}

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

// ── PPTX effects bridge ──

import type {
  EffectListOptions,
  Scene3DOptions,
  Shape3DOptions,
} from "@office-open/core/drawingml";

function toColor(color?: string, alpha?: number) {
  if (!color) return { value: "000000", alpha: (alpha ?? 40) * 1000 };
  return { value: color.replace("#", ""), alpha: (alpha ?? 40) * 1000 };
}

/** Map PPTX EffectsOptions to core EffectListOptions. */
function toEffectListOptions(opts: EffectsOptions): EffectListOptions | undefined {
  const hasEffects =
    opts.outerShadow || opts.innerShadow || opts.glow || opts.reflection || opts.softEdge;
  if (!hasEffects) return undefined;

  return {
    outerShadow: opts.outerShadow
      ? {
          blurRadius: opts.outerShadow.blur,
          distance: opts.outerShadow.distance,
          direction: opts.outerShadow.direction,
          rotWithShape: opts.outerShadow.rotateWithShape === false ? false : undefined,
          color: toColor(opts.outerShadow.color, opts.outerShadow.alpha),
        }
      : undefined,
    innerShadow: opts.innerShadow
      ? {
          blurRadius: opts.innerShadow.blur,
          distance: opts.innerShadow.distance,
          direction: opts.innerShadow.direction,
          color: toColor(opts.innerShadow.color, opts.innerShadow.alpha),
        }
      : undefined,
    glow: opts.glow
      ? {
          radius: opts.glow.radius ?? 152400,
          color: toColor(opts.glow.color, opts.glow.alpha),
        }
      : undefined,
    reflection: opts.reflection ? toReflectionCore(opts.reflection) : undefined,
    softEdge: opts.softEdge ? (opts.softEdge.radius ?? 50800) : undefined,
  };
}

function toReflectionCore(opts: ReflectionOptions) {
  const result: Record<string, number | string> = {};
  if (opts.blurRadius !== undefined) result.blurRadius = opts.blurRadius;
  if (opts.distance !== undefined) result.distance = opts.distance;
  if (opts.direction !== undefined) result.direction = opts.direction;
  if (opts.startAlpha !== undefined) result.startAlpha = opts.startAlpha * 1000;
  if (opts.startPosition !== undefined) result.startPosition = opts.startPosition * 1000;
  if (opts.endAlpha !== undefined) result.endAlpha = opts.endAlpha * 1000;
  if (opts.endPosition !== undefined) result.endPosition = opts.endPosition * 1000;
  if (opts.fadeDirection !== undefined) result.fadeDirection = opts.fadeDirection * 60000;
  if (opts.scaleX !== undefined) result.scaleX = opts.scaleX * 1000;
  if (opts.scaleY !== undefined) result.scaleY = opts.scaleY * 1000;
  if (opts.skewX !== undefined) result.skewX = opts.skewX * 60000;
  if (opts.skewY !== undefined) result.skewY = opts.skewY * 60000;
  if (opts.alignment !== undefined) result.alignment = opts.alignment;
  if (opts.rotateWithShape === false) result.rotWithShape = 0;
  return result;
}

/** Map PPTX EffectsOptions to core Scene3DOptions. */
function toScene3DOptions(opts: EffectsOptions): Scene3DOptions | undefined {
  if (!opts.rotation3D && !opts.lighting) return undefined;

  const cameraPreset = opts.rotation3D?.perspective
    ? "legacyPerspectiveFront"
    : "orthographicFront";
  const cameraOpts = {
    preset: cameraPreset,
    ...(opts.rotation3D?.perspective && { fov: opts.rotation3D.perspective }),
    ...(opts.rotation3D && {
      rotation: {
        lat: (opts.rotation3D.x ?? 0) * 60000,
        lon: (opts.rotation3D.y ?? 0) * 60000,
        rev: (opts.rotation3D.z ?? 0) * 60000,
      },
    }),
  };

  return {
    camera: cameraOpts as Scene3DOptions["camera"],
    lightRig: { rig: opts.lighting ?? "threePt", direction: "t" },
  };
}

/** Map PPTX EffectsOptions to core Shape3DOptions. */
function toShape3DOptions(opts: EffectsOptions): Shape3DOptions | undefined {
  if (!opts.extrusionH && !opts.bevelTop && !opts.bevelBottom && !opts.material) return undefined;

  return {
    ...(opts.bevelTop
      ? { bevelT: { w: opts.bevelTop.width! * 12700, h: opts.bevelTop.height! * 12700 } }
      : {}),
    ...(opts.bevelBottom
      ? { bevelB: { w: opts.bevelBottom.width! * 12700, h: opts.bevelBottom.height! * 12700 } }
      : {}),
    ...(opts.extrusionH !== undefined ? { extrusionH: opts.extrusionH } : {}),
    ...(opts.material ? { prstMaterial: opts.material } : {}),
  };
}

// ── PPTX outline bridge ──

const DASH_STYLE_MAP = {
  solid: "solid",
  dash: "dash",
  dashDot: "dashDot",
  lgDash: "lgDash",
  sysDot: "sysDot",
  sysDash: "sysDash",
} as const;

export function toCoreOutlineOptions(opts: OutlineOptions): CoreOutlineOptions {
  const result: CoreOutlineOptions = {
    width: opts.width,
  };
  if (opts.color) {
    result.type = "solidFill";
    result.color = { value: opts.color.replace("#", "") };
  } else {
    result.type = "noFill";
  }
  if (opts.dashStyle) {
    result.dash =
      (
        DASH_STYLE_MAP as Record<
          string,
          "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash"
        >
      )[opts.dashStyle] ?? "solid";
  }
  return result;
}

// ── Shape (p:sp) descriptor ──

export const shapeDesc: CustomDescriptor<ShapeDescriptorOptions> = {
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
      const styleXml = stringifyStyle(opts.style);
      if (styleXml) parts.push(styleXml);
    }

    // ── p:txBody ──
    parts.push(stringifyTxBody(opts.textBody ?? {}, ctx));

    // ── Root attributes ──
    const spAttrs: string[] = [];
    if (opts.useBackgroundFill) spAttrs.push(' useBgFill="1"');
    if (opts.blackWhiteMode) spAttrs.push(` bwMode="${opts.blackWhiteMode}"`);

    return `<p:sp${spAttrs.join("")}>${parts.join("")}</p:sp>`;
  },

  parse(el, ctx) {
    const result: ShapeDescriptorOptions = {};

    // Root attributes
    if (el.attributes) {
      if (el.attributes["useBgFill"] !== undefined)
        result.useBackgroundFill = el.attributes["useBgFill"] === "1";
      if (el.attributes["bwMode"] !== undefined)
        result.blackWhiteMode = String(
          el.attributes["bwMode"],
        ) as ShapeDescriptorOptions["blackWhiteMode"];
    }

    // p:nvSpPr
    const nvSpPr = findChild(el, "p:nvSpPr");
    if (nvSpPr) {
      const parsed = readNvSpPr(nvSpPr);
      if (parsed.id !== undefined) result.id = parsed.id;
      if (parsed.name !== undefined) result.name = parsed.name;
      if (parsed.placeholder !== undefined) result.placeholder = parsed.placeholder;
      if (parsed.placeholderIndex !== undefined) result.placeholderIndex = parsed.placeholderIndex;
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
    if (style) result.style = readStyle(style);

    // p:txBody
    const txBody = findChild(el, "p:txBody");
    if (txBody) result.textBody = readTxBody(txBody, ctx);

    return result;
  },
};

// ── Picture (p:pic) descriptor ──

export const pictureDesc: CustomDescriptor<PictureDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextPictureId++;
    const name = opts.name ?? `Picture ${id}`;
    const fileName = `${name.replace(/\s+/g, "_")}.${opts.type}`;

    // Calculate EMU values for width/height
    const widthEmu = toEmu(opts.width) ?? 0;
    const heightEmu = toEmu(opts.height) ?? 0;
    const widthPixels = typeof opts.width === "number" ? opts.width : 0;
    const heightPixels = typeof opts.height === "number" ? opts.height : 0;

    // Register media with the PPTX context
    pptx.addImage(fileName, {
      key: fileName,
      type: opts.type,
      fileName,
      data: opts.data,
      transformation: {
        pixels: { x: widthPixels, y: heightPixels },
        emus: { x: widthEmu, y: heightEmu },
      },
    });

    const parts: string[] = [];

    // ── p:nvPicPr ──
    parts.push(stringifyNvPicPr(id, name));

    // ── p:blipFill ──
    parts.push(stringifyPptxBlipFill(fileName));

    // ── p:spPr ──
    const spPrXml = stringifyPicSpPr(opts, ctx);
    if (spPrXml) parts.push(spPrXml);

    return `<p:pic>${parts.join("")}</p:pic>`;
  },

  parse(el, ctx) {
    const result: Partial<PictureDescriptorOptions> = {};

    // p:nvPicPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "p:cNvPr");
      if (cNvPr?.attributes) {
        if (cNvPr.attributes["id"] !== undefined) result.id = Number(cNvPr.attributes["id"]);
        if (cNvPr.attributes["name"] !== undefined) result.name = String(cNvPr.attributes["name"]);
      }
    }

    // p:spPr (position/size only)
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) {
        const off = findChild(xfrm, "a:off");
        if (off?.attributes) {
          result.x = Math.round(Number(off.attributes["x"] ?? 0) / 9525);
          result.y = Math.round(Number(off.attributes["y"] ?? 0) / 9525);
        }
        const ext = findChild(xfrm, "a:ext");
        if (ext?.attributes) {
          result.width = Math.round(Number(ext.attributes["cx"] ?? 0) / 9525);
          result.height = Math.round(Number(ext.attributes["cy"] ?? 0) / 9525);
        }
      }
    }

    // Image data from p:blipFill → a:blip → r:embed
    const blip = findDeep(el, "a:blip")[0];
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

    return result as PictureDescriptorOptions;
  },
};

// ── Shape helper: p:nvSpPr ──

function stringifyNvSpPr(id: number, name: string, opts: ShapeDescriptorOptions): string {
  // nvPr
  let nvPrContent = "<p:nvPr/>";
  if (opts.placeholder) {
    const phAttrs: string[] = [`type="${opts.placeholder}"`];
    if (opts.placeholderIndex !== undefined) phAttrs.push(`idx="${opts.placeholderIndex}"`);
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

  return `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>${cNvSpPrContent}${nvPrContent}</p:nvSpPr>`;
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

/** Convert PPTX dimension (pixels or UniversalMeasure) to EMU. */
function toEmu(val: number | string | undefined): number | undefined {
  if (val === undefined) return undefined;
  // PPTX numbers are pixels, strings are UniversalMeasure
  return typeof val === "string" ? convertToEmu(val) : convertPixelsToEmu(val);
}

function stringifySpPr(opts: ShapeDescriptorOptions, ctx: WriteContext): string {
  const pptx = ctx as PptxWriteContext;
  const parts: string[] = [];

  // Transform2D
  const hasTransform =
    opts.x !== undefined ||
    opts.y !== undefined ||
    opts.width !== undefined ||
    opts.height !== undefined ||
    opts.flipHorizontal !== undefined ||
    opts.rotation !== undefined;

  if (hasTransform) {
    const transformOpts: Transform2DOptions = {
      x: toEmu(opts.x),
      y: toEmu(opts.y),
      width: toEmu(opts.width),
      height: toEmu(opts.height),
      flipHorizontal: opts.flipHorizontal,
      rotation: opts.rotation,
    };
    const xml = stringify(transform2DDesc, transformOpts, ctx);
    if (xml) parts.push(xml);
  }

  // PresetGeometry
  const geomXml = stringify(presetGeometryDesc, { preset: opts.geometry ?? "rect" }, ctx);
  if (geomXml) parts.push(geomXml);

  // Fill (with blipFill media registration side effect)
  if (opts.fill) {
    const fillType = typeof opts.fill === "string" ? "solid" : opts.fill.type;
    if (fillType === "blip") {
      const blipFill = opts.fill as Extract<CoreFillOptions, { type: "blip" }>;
      if (blipFill.data) {
        const raw =
          blipFill.data instanceof Uint8Array
            ? blipFill.data
            : new TextEncoder().encode(blipFill.data);
        const fileName = `image_blip.${blipFill.imageType ?? "png"}`;
        pptx.addImage(fileName, {
          key: fileName,
          data: raw,
          fileName,
          type: (blipFill.imageType ?? "png") as MediaEntry["type"],
          transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
        });
      }
    }
    const fillXml = stringify(fillDesc, opts.fill, ctx);
    if (fillXml) parts.push(fillXml);
  } else {
    parts.push("<a:noFill/>");
  }

  // Outline
  if (opts.outline) {
    const coreOutline = toCoreOutlineOptions(opts.outline);
    const outlineXml = stringify(outlineDesc, coreOutline, ctx);
    if (outlineXml) parts.push(outlineXml);
  }

  // Effects
  if (opts.effects) {
    const effectListOpts = toEffectListOptions(opts.effects);
    if (effectListOpts) {
      const xml = stringify(effectListDesc, effectListOpts, ctx);
      if (xml) parts.push(xml);
    }

    const scene3DOpts = toScene3DOptions(opts.effects);
    if (scene3DOpts) {
      const xml = stringify(scene3DDesc, scene3DOpts, ctx);
      if (xml) parts.push(xml);
    }

    const shape3DOpts = toShape3DOptions(opts.effects);
    if (shape3DOpts) {
      const xml = stringify(shape3DDesc, shape3DOpts, ctx);
      if (xml) parts.push(xml);
    }
  }

  if (parts.length === 0) return "<p:spPr/>";
  return `<p:spPr>${parts.join("")}</p:spPr>`;
}

// ── Shape helper: p:style ──

function stringifyStyle(style: ShapeStyleDescriptorOptions): string | undefined {
  const parts: string[] = [];

  if (style.lineReference) {
    const colorChild = style.lineReference.color
      ? `<a:srgbClr val="${style.lineReference.color}"/>`
      : "";
    parts.push(`<a:lnRef idx="${style.lineReference.index}">${colorChild}</a:lnRef>`);
  }
  if (style.fillReference) {
    const colorChild = style.fillReference.color
      ? `<a:srgbClr val="${style.fillReference.color}"/>`
      : "";
    parts.push(`<a:fillRef idx="${style.fillReference.index}">${colorChild}</a:fillRef>`);
  }
  if (style.effectReference) {
    const colorChild = style.effectReference.color
      ? `<a:srgbClr val="${style.effectReference.color}"/>`
      : "";
    parts.push(`<a:effectRef idx="${style.effectReference.index}">${colorChild}</a:effectRef>`);
  }
  if (style.fontReference) {
    const colorChild = style.fontReference.color
      ? `<a:solidFill><a:srgbClr val="${style.fontReference.color}"/></a:solidFill>`
      : "";
    parts.push(`<a:fontRef idx="${style.fontReference.index}">${colorChild}</a:fontRef>`);
  }

  if (parts.length === 0) return undefined;
  return `<p:style>${parts.join("")}</p:style>`;
}

// ── Shape helper: p:txBody ──

function stringifyTxBody(opts: TextBodyDescriptorOptions, ctx: WriteContext): string {
  const parts: string[] = [];

  // a:bodyPr
  parts.push(stringifyBodyPr(opts));

  // a:lstStyle
  parts.push("<a:lstStyle/>");

  // Paragraphs — use paragraphDesc from text.ts
  if (opts.children && opts.children.length > 0) {
    for (const p of opts.children) {
      if (typeof p === "string") {
        const xml = paragraphDesc.stringify({ text: p }, ctx);
        parts.push(xml ?? "<a:p/>");
      } else {
        const xml = paragraphDesc.stringify(toParagraphDescOpts(p), ctx);
        parts.push(xml ?? "<a:p/>");
      }
    }
  } else if (opts.text !== undefined) {
    const xml = paragraphDesc.stringify({ text: opts.text }, ctx);
    parts.push(xml ?? "<a:p/>");
  } else {
    parts.push("<a:p/>");
  }

  return `<p:txBody>${parts.join("")}</p:txBody>`;
}

/** Convert ParagraphOptions to ParagraphDescriptorOptions. */
function toParagraphDescOpts(opts: ParagraphOptions): ParagraphDescriptorOptions {
  const result: ParagraphDescriptorOptions = {};
  if (opts.text !== undefined) result.text = opts.text;
  if (opts.children !== undefined) result.children = opts.children as any;
  if (opts.properties) result.properties = opts.properties;
  return result;
}

function stringifyBodyPr(opts: TextBodyDescriptorOptions): string {
  const bodyPrChildren: string[] = [];

  if (opts.autoFit === "normal") bodyPrChildren.push("<a:normAutofit/>");
  else if (opts.autoFit === "shape") bodyPrChildren.push("<a:spAutoFit/>");
  else if (opts.autoFit === "none") bodyPrChildren.push("<a:noAutofit/>");

  const attrs: string[] = [];
  if (opts.vertical) attrs.push(`vert="${opts.vertical}"`);
  if (opts.anchor) attrs.push(`anchor="${xsdTextAnchor.to(opts.anchor)}"`);
  if (opts.wrap) attrs.push(`wrap="${opts.wrap}"`);
  if (opts.margins?.top !== undefined) attrs.push(`tIns="${opts.margins.top}"`);
  if (opts.margins?.bottom !== undefined) attrs.push(`bIns="${opts.margins.bottom}"`);
  if (opts.margins?.left !== undefined) attrs.push(`lIns="${opts.margins.left}"`);
  if (opts.margins?.right !== undefined) attrs.push(`rIns="${opts.margins.right}"`);
  if (opts.columns !== undefined) attrs.push(`numCol="${opts.columns}"`);
  if (opts.columnSpacing !== undefined) attrs.push(`spcCol="${opts.columnSpacing * 100}"`);
  if (opts.marginTop !== undefined) attrs.push(`marT="${opts.marginTop}"`);
  if (opts.marginBottom !== undefined) attrs.push(`marB="${opts.marginBottom}"`);

  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  const children = bodyPrChildren.join("");

  if (!children && !attrStr) return "<a:bodyPr/>";
  if (!children) return `<a:bodyPr${attrStr}/>`;
  return `<a:bodyPr${attrStr}>${children}</a:bodyPr>`;
}

// ── Picture helpers ──

function stringifyNvPicPr(id: number, name: string): string {
  return `<p:nvPicPr><p:cNvPr id="${id}" name="${escapeXml(name)}" descr=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`;
}

/** PPTX uses p:blipFill (not pic:blipFill). */
function stringifyPptxBlipFill(fileName: string): string {
  return `<p:blipFill><a:blip r:embed="{image:${escapeXml(fileName)}}" cstate="none"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
}

function stringifyPicSpPr(opts: PictureDescriptorOptions, ctx: WriteContext): string {
  const parts: string[] = [];

  // Transform2D (always present for pictures)
  const transformOpts: Transform2DOptions = {
    x: toEmu(opts.x),
    y: toEmu(opts.y),
    width: toEmu(opts.width),
    height: toEmu(opts.height),
  };
  const xfrmXml = stringify(transform2DDesc, transformOpts, ctx);
  if (xfrmXml) parts.push(xfrmXml);

  // PresetGeometry (always rect for pictures)
  const geomXml = stringify(presetGeometryDesc, { preset: "rect" }, ctx);
  if (geomXml) parts.push(geomXml);

  if (parts.length === 0) return "<p:spPr/>";
  return `<p:spPr>${parts.join("")}</p:spPr>`;
}

// ── Read helpers ──

export function readNvSpPr(nvSpPr: XmlElement): ShapeDescriptorOptions {
  const result: ShapeDescriptorOptions = {};

  const cNvPr = findChild(nvSpPr, "p:cNvPr");
  if (cNvPr?.attributes) {
    if (cNvPr.attributes["id"] !== undefined) result.id = Number(cNvPr.attributes["id"]);
    if (cNvPr.attributes["name"] !== undefined) result.name = String(cNvPr.attributes["name"]);
  }

  const nvPr = findChild(nvSpPr, "p:nvPr");
  if (nvPr) {
    if (nvPr.attributes) {
      if (nvPr.attributes["isPhoto"] !== undefined)
        result.isPhoto = nvPr.attributes["isPhoto"] === "1";
      if (nvPr.attributes["userDrawn"] !== undefined)
        result.userDrawn = nvPr.attributes["userDrawn"] === "1";
    }
    const ph = findChild(nvPr, "p:ph");
    if (ph?.attributes) {
      if (ph.attributes["type"] !== undefined)
        result.placeholder = String(ph.attributes["type"]) as ShapeDescriptorOptions["placeholder"];
      if (ph.attributes["idx"] !== undefined)
        result.placeholderIndex = Number(ph.attributes["idx"]);
      if (ph.attributes["hasCustomPrompt"] !== undefined)
        result.hasCustomPrompt = ph.attributes["hasCustomPrompt"] === "1";
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

export function readSpPr(spPr: XmlElement, ctx: ReadContext): ShapeDescriptorOptions {
  const result: ShapeDescriptorOptions = {};

  // Transform
  const xfrm = findChild(spPr, "a:xfrm");
  if (xfrm) {
    const transformOpts = parse(transform2DDesc, xfrm, ctx);
    if (transformOpts.x !== undefined) result.x = Math.round(convertToEmu(transformOpts.x) / 9525);
    if (transformOpts.y !== undefined) result.y = Math.round(convertToEmu(transformOpts.y) / 9525);
    if (transformOpts.width !== undefined)
      result.width = Math.round(convertToEmu(transformOpts.width) / 9525);
    if (transformOpts.height !== undefined)
      result.height = Math.round(convertToEmu(transformOpts.height) / 9525);
    if (transformOpts.flipHorizontal !== undefined)
      result.flipHorizontal = transformOpts.flipHorizontal;
    if (transformOpts.rotation !== undefined) result.rotation = transformOpts.rotation;
  }

  // Geometry
  const prstGeom = findChild(spPr, "a:prstGeom");
  if (prstGeom) {
    const geomOpts = parse(presetGeometryDesc, prstGeom, ctx);
    if (geomOpts.preset) result.geometry = geomOpts.preset;
  }

  // Fill
  const fillResult = parse(fillDesc, spPr, ctx);
  if (fillResult && Object.keys(fillResult).length > 0) {
    result.fill = fillResult as ShapeDescriptorOptions["fill"];
  }

  // Outline
  const ln = findChild(spPr, "a:ln");
  if (ln) {
    result.outline = readOutlineCompat(ln);
  }

  // Effects
  const effects = readEffectsFromSpPr(spPr);
  if (effects) result.effects = effects as ShapeDescriptorOptions["effects"];

  return result;
}

/** Read x/y/width/height (in pixels) from an a:xfrm element. */
export function readPositionFromXfrm(xfrm: XmlElement): Record<string, number> {
  const result: Record<string, number> = {};
  const off = findChild(xfrm, "a:off");
  if (off) {
    const x = attrNum(off, "x");
    if (x !== undefined) result.x = Math.round(x / 9525);
    const y = attrNum(off, "y");
    if (y !== undefined) result.y = Math.round(y / 9525);
  }
  const ext = findChild(xfrm, "a:ext");
  if (ext) {
    const cx = attrNum(ext, "cx");
    if (cx !== undefined) result.width = Math.round(cx / 9525);
    const cy = attrNum(ext, "cy");
    if (cy !== undefined) result.height = Math.round(cy / 9525);
  }
  return result;
}

export function readOutlineCompat(ln: XmlElement): OutlineOptions {
  const coreOpts = parse(outlineDesc, ln, {} as ReadContext);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const result: Record<string, any> = {};

  if (coreOpts.width !== undefined) result.width = coreOpts.width;
  if (coreOpts.dash) result.dashStyle = coreOpts.dash;

  if (coreOpts.type === "solidFill" && coreOpts.color) {
    const c = coreOpts.color as { value?: string };
    result.color = c.value ?? "";
  }

  return result as OutlineOptions;
}

function readStyle(styleEl: XmlElement): ShapeStyleDescriptorOptions {
  const result: ShapeStyleDescriptorOptions = {};

  const lnRef = findChild(styleEl, "a:lnRef");
  if (lnRef?.attributes) {
    result.lineReference = { index: Number(lnRef.attributes["idx"] ?? 0) };
    const srgbClr = findChild(lnRef, "a:srgbClr");
    if (srgbClr?.attributes?.["val"])
      result.lineReference.color = String(srgbClr.attributes["val"]);
  }

  const fillRef = findChild(styleEl, "a:fillRef");
  if (fillRef?.attributes) {
    result.fillReference = { index: Number(fillRef.attributes["idx"] ?? 0) };
    const srgbClr = findChild(fillRef, "a:srgbClr");
    if (srgbClr?.attributes?.["val"])
      result.fillReference.color = String(srgbClr.attributes["val"]);
  }

  const effectRef = findChild(styleEl, "a:effectRef");
  if (effectRef?.attributes) {
    result.effectReference = { index: Number(effectRef.attributes["idx"] ?? 0) };
    const srgbClr = findChild(effectRef, "a:srgbClr");
    if (srgbClr?.attributes?.["val"])
      result.effectReference.color = String(srgbClr.attributes["val"]);
  }

  const fontRef = findChild(styleEl, "a:fontRef");
  if (fontRef?.attributes) {
    result.fontReference = { index: Number(fontRef.attributes["idx"] ?? 0) };
    const solidFill = findChild(fontRef, "a:solidFill");
    if (solidFill) {
      const srgbClr = findChild(solidFill, "a:srgbClr");
      if (srgbClr?.attributes?.["val"])
        result.fontReference.color = String(srgbClr.attributes["val"]);
    }
  }

  return result;
}

function readTxBody(txBody: XmlElement, ctx: ReadContext): TextBodyDescriptorOptions {
  const result: TextBodyDescriptorOptions = {};

  // a:bodyPr
  const bodyPr = findChild(txBody, "a:bodyPr");
  if (bodyPr) {
    const attrs = bodyPr.attributes ?? {};
    if (attrs["vert"] !== undefined)
      result.vertical = String(attrs["vert"]) as TextBodyDescriptorOptions["vertical"];
    if (attrs["anchor"] !== undefined)
      result.anchor = xsdTextAnchor.from(
        String(attrs["anchor"]),
      ) as TextBodyDescriptorOptions["anchor"];
    if (attrs["wrap"] !== undefined)
      result.wrap = String(attrs["wrap"]) as TextBodyDescriptorOptions["wrap"];
    if (attrs["tIns"] !== undefined)
      result.margins = { ...result.margins, top: Number(attrs["tIns"]) };
    if (attrs["bIns"] !== undefined)
      result.margins = { ...result.margins, bottom: Number(attrs["bIns"]) };
    if (attrs["lIns"] !== undefined)
      result.margins = { ...result.margins, left: Number(attrs["lIns"]) };
    if (attrs["rIns"] !== undefined)
      result.margins = { ...result.margins, right: Number(attrs["rIns"]) };
    if (attrs["numCol"] !== undefined) result.columns = Number(attrs["numCol"]);
    if (attrs["spcCol"] !== undefined) result.columnSpacing = Number(attrs["spcCol"]) / 100;
    if (attrs["marT"] !== undefined) result.marginTop = Number(attrs["marT"]);
    if (attrs["marB"] !== undefined) result.marginBottom = Number(attrs["marB"]);

    if (findChild(bodyPr, "a:normAutofit")) result.autoFit = "normal";
    else if (findChild(bodyPr, "a:spAutoFit")) result.autoFit = "shape";
    else if (findChild(bodyPr, "a:noAutofit")) result.autoFit = "none";
  }

  // Paragraphs — use paragraphDesc
  const paragraphs: ParagraphDescriptorOptions[] = [];
  if (txBody.elements) {
    for (const child of txBody.elements) {
      if (child.name === "a:p") {
        const para = paragraphDesc.parse(child, ctx);
        paragraphs.push(para as ParagraphDescriptorOptions);
      }
    }
  }

  if (paragraphs.length === 1 && paragraphs[0].text && !paragraphs[0].properties) {
    result.text = paragraphs[0].text;
  } else if (paragraphs.length > 0) {
    result.children = paragraphs.map((p) => {
      if (p.text) return p.text;
      return p as ParagraphOptions;
    });
  }

  return result;
}

// ── Image type helper ──

/** Map file extension to PictureOptions type. */
function imageTypeFromPath(path: string): "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf" {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";
  switch (ext) {
    case "jpg":
    case "jpeg":
      return "jpg";
    case "png":
      return "png";
    case "gif":
      return "gif";
    case "bmp":
      return "bmp";
    case "emf":
      return "emf";
    case "wmf":
      return "wmf";
    default:
      return "png";
  }
}

// ── Effects parsing ──

/** Extract color and alpha from an effect element's child color element. */
function extractColorFromElement(el: XmlElement): { color?: string; alpha?: number } | undefined {
  for (const child of el.elements ?? []) {
    if (child.name === "a:srgbClr") {
      const color = child.attributes?.["val"] ? String(child.attributes["val"]) : undefined;
      let alpha: number | undefined;
      for (const transform of child.elements ?? []) {
        if (transform.name === "a:alpha") {
          const val = attrNum(transform, "val");
          if (val !== undefined) alpha = Math.round(val / 1000);
        }
      }
      return { color, alpha };
    }
  }
  return undefined;
}

/** Parse 2D effects (a:effectLst) into PPTX EffectsOptions fields. */
export function readEffectList(effectLst: XmlElement): Record<string, unknown> | undefined {
  const opts: Record<string, unknown> = {};
  for (const child of effectLst.elements ?? []) {
    if (!child.name) continue;
    switch (child.name) {
      case "a:outerShdw": {
        const shadow: Record<string, unknown> = {};
        const blurRad = attrNum(child, "blurRad");
        if (blurRad !== undefined) shadow.blur = blurRad;
        const dist = attrNum(child, "dist");
        if (dist !== undefined) shadow.distance = dist;
        const dir = attrNum(child, "dir");
        if (dir !== undefined) shadow.direction = dir;
        const color = extractColorFromElement(child);
        if (color) {
          if (color.color) shadow.color = color.color;
          if (color.alpha !== undefined) shadow.alpha = color.alpha;
        }
        opts.outerShadow = shadow;
        break;
      }
      case "a:innerShdw": {
        const shadow: Record<string, unknown> = {};
        const blurRad = attrNum(child, "blurRad");
        if (blurRad !== undefined) shadow.blur = blurRad;
        const dist = attrNum(child, "dist");
        if (dist !== undefined) shadow.distance = dist;
        const dir = attrNum(child, "dir");
        if (dir !== undefined) shadow.direction = dir;
        const color = extractColorFromElement(child);
        if (color) {
          if (color.color) shadow.color = color.color;
          if (color.alpha !== undefined) shadow.alpha = color.alpha;
        }
        opts.innerShadow = shadow;
        break;
      }
      case "a:glow": {
        const glow: Record<string, unknown> = {};
        const rad = attrNum(child, "rad");
        if (rad !== undefined) glow.radius = rad;
        const color = extractColorFromElement(child);
        if (color) {
          if (color.color) glow.color = color.color;
          if (color.alpha !== undefined) glow.alpha = color.alpha;
        }
        opts.glow = glow;
        break;
      }
      case "a:reflection": {
        // CT_ReflectionEffect has 14 attrs; toReflectionCore writes them all
        // (with unit scaling). Invert each scale on read.
        const reflection: ReflectionOptions = {};
        const blurRad = attrNum(child, "blurRad");
        if (blurRad !== undefined) reflection.blurRadius = blurRad;
        const dist = attrNum(child, "dist");
        if (dist !== undefined) reflection.distance = dist;
        const dir = attrNum(child, "dir");
        if (dir !== undefined) reflection.direction = dir;
        const stA = attrNum(child, "stA");
        if (stA !== undefined) reflection.startAlpha = stA / 1000;
        const stPos = attrNum(child, "stPos");
        if (stPos !== undefined) reflection.startPosition = stPos / 1000;
        const endA = attrNum(child, "endA");
        if (endA !== undefined) reflection.endAlpha = endA / 1000;
        const endPos = attrNum(child, "endPos");
        if (endPos !== undefined) reflection.endPosition = endPos / 1000;
        const fadeDir = attrNum(child, "fadeDir");
        if (fadeDir !== undefined) reflection.fadeDirection = fadeDir / 60000;
        const sx = attrNum(child, "sx");
        if (sx !== undefined) reflection.scaleX = sx / 1000;
        const sy = attrNum(child, "sy");
        if (sy !== undefined) reflection.scaleY = sy / 1000;
        const kx = attrNum(child, "kx");
        if (kx !== undefined) reflection.skewX = kx / 60000;
        const ky = attrNum(child, "ky");
        if (ky !== undefined) reflection.skewY = ky / 60000;
        const algn = attr(child, "algn");
        if (algn)
          reflection.alignment = xsdRectAlignment.from(algn) as ReflectionOptions["alignment"];
        const rotWithShape = attr(child, "rotWithShape");
        if (rotWithShape !== undefined) reflection.rotateWithShape = rotWithShape !== "0";
        opts.reflection = reflection;
        break;
      }
      case "a:softEdge": {
        const rad = attrNum(child, "rad");
        if (rad !== undefined) opts.softEdge = { radius: rad };
        break;
      }
    }
  }
  return Object.keys(opts).length > 0 ? opts : undefined;
}

/** Parse effects from p:spPr (2D effects, 3D scene, 3D shape props). */
function readEffectsFromSpPr(spPr: XmlElement): Record<string, unknown> | undefined {
  const opts: Record<string, unknown> = {};

  // 2D effects from a:effectLst
  const effectLst = findChild(spPr, "a:effectLst");
  if (effectLst) {
    const effectOpts = readEffectList(effectLst);
    if (effectOpts) Object.assign(opts, effectOpts);
  }

  // 3D scene (rotation3D, lighting) from a:scene3d
  const scene3d = findChild(spPr, "a:scene3d");
  if (scene3d) {
    const camera = findChild(scene3d, "a:camera");
    if (camera) {
      const rot = findChild(camera, "a:rot");
      if (rot) {
        const rotation3d: Record<string, number> = {};
        const lat = attrNum(rot, "lat");
        const lon = attrNum(rot, "lon");
        const rev = attrNum(rot, "rev");
        if (lat !== undefined) rotation3d.x = Math.round(lat / 60000);
        if (lon !== undefined) rotation3d.y = Math.round(lon / 60000);
        if (rev !== undefined) rotation3d.z = Math.round(rev / 60000);
        const fov = attrNum(camera, "fov");
        if (fov !== undefined) rotation3d.perspective = fov;
        if (Object.keys(rotation3d).length > 0) opts.rotation3D = rotation3d;
      }
    }
    const lightRig = findChild(scene3d, "a:lightRig");
    if (lightRig) {
      const rig = attr(lightRig, "rig");
      if (rig) opts.lighting = rig;
    }
  }

  // 3D shape properties (bevel, extrusion, material) from a:sp3d
  const sp3d = findChild(spPr, "a:sp3d");
  if (sp3d) {
    const bevelT = findChild(sp3d, "a:bevelT");
    if (bevelT) {
      const bevel: Record<string, unknown> = {};
      const w = attrNum(bevelT, "w");
      const h = attrNum(bevelT, "h");
      if (w !== undefined) bevel.width = Math.round(w / 12700);
      if (h !== undefined) bevel.height = Math.round(h / 12700);
      const prst = attr(bevelT, "prst");
      if (prst) bevel.preset = prst;
      if (Object.keys(bevel).length > 0) opts.bevelTop = bevel;
    }
    const bevelB = findChild(sp3d, "a:bevelB");
    if (bevelB) {
      const bevel: Record<string, unknown> = {};
      const w = attrNum(bevelB, "w");
      const h = attrNum(bevelB, "h");
      if (w !== undefined) bevel.width = Math.round(w / 12700);
      if (h !== undefined) bevel.height = Math.round(h / 12700);
      const prst = attr(bevelB, "prst");
      if (prst) bevel.preset = prst;
      if (Object.keys(bevel).length > 0) opts.bevelBottom = bevel;
    }
    const z = attrNum(sp3d, "z");
    if (z !== undefined) opts.depth = z;
    const contourW = attrNum(sp3d, "contourW");
    if (contourW !== undefined) opts.contourWidth = contourW;
    const extrusionH = attrNum(sp3d, "extrusionH");
    if (extrusionH !== undefined) opts.extrusionH = extrusionH;
    const prstMaterial = attr(sp3d, "prstMaterial");
    if (prstMaterial) opts.material = prstMaterial;
  }

  return Object.keys(opts).length > 0 ? opts : undefined;
}
