/**
 * Text body properties for DrawingML shapes (CT_TextBodyProperties).
 *
 * Defines how text is laid out within a shape: margins, alignment, autofit,
 * vertical text, overflow, columns, and 3D. The factory + parser cover all 19
 * CT_TextBodyProperties attributes and child elements (prstTxWarp, EG_TextAutofit,
 * scene3d, EG_Text3D).
 *
 * The container tag is parameterized: PPTX/XLSX emit a:bodyPr, DOCX emits
 * wps:bodyPr. `createBodyProperties` takes an optional tagName (default a:bodyPr);
 * `bodyPropertiesDesc` is the a:bodyPr descriptor instance.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TextBodyProperties
 *
 * @module
 */

import { attr, attrBool, attrMeasure, attrNum, element, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ReadContext, WriteContext, CustomDescriptor } from "../../descriptor";
import { convertToEmu } from "../../util/converters";
import type { UniversalMeasure } from "../../util/values";
import { createScene3D } from "../three-d/scene-3d";
import type { Scene3DOptions } from "../three-d/scene-3d";
import { createShape3D } from "../three-d/shape-3d";
import type { Shape3DOptions } from "../three-d/shape-3d";
import { scene3DDesc, shape3DDesc } from "../three-d/three-d-descriptors";

// ── Enumerations ──

/** Text anchoring type (ST_TextAnchoringType). */
export enum VerticalAnchor {
  TOP = "t",
  CENTER = "ctr",
  BOTTOM = "b",
  JUSTIFY = "just",
  DISTRIBUTED = "dist",
}

/** Text vertical overflow type (ST_TextVertOverflowType). */
export const TextVertOverflowType = {
  OVERFLOW: "overflow",
  ELLIPSIS: "ellipsis",
  CLIP: "clip",
} as const;

/** Text horizontal overflow type (ST_TextHorzOverflowType). */
export const TextHorzOverflowType = {
  OVERFLOW: "overflow",
  CLIP: "clip",
} as const;

/** Text vertical type (ST_TextVerticalType). */
export const TextVerticalType = {
  HORIZONTAL: "horz",
  VERTICAL: "vert",
  VERTICAL_270: "vert270",
  WORD_ART_VERTICAL: "wordArtVert",
  EAST_ASIAN_VERTICAL: "eaVert",
  MONGOLIAN_VERTICAL: "mongolianVert",
  WORD_ART_VERTICAL_RTL: "wordArtVertRtl",
} as const;

/** Text body wrapping type (ST_TextWrappingType). */
export const TextBodyWrappingType = {
  NONE: "none",
  SQUARE: "square",
} as const;

// ── Sub-element Options ──

/** Normal autofit options (CT_TextNormalAutofit). */
export interface NormalAutofitOptions {
  /** Font scale percentage (e.g., 100000 = 100%). */
  fontScale?: number;
  /** Line spacing reduction percentage. */
  lnSpcReduction?: number;
}

/** Preset text warp options (CT_PresetTextShape). */
export interface PresetTextShapeOptions {
  /** Preset shape type (e.g. "textArchUp", "textCircle"). */
  preset: string;
  /** Optional adjustment values. */
  adjustments?: { name: string; formula: string }[];
}

/** Flat text options (CT_FlatText). */
export interface FlatTextOptions {
  /** Z-offset in EMUs. */
  z?: number;
}

// ── Main Options ──

/**
 * Options for text body properties (CT_TextBodyProperties).
 *
 * XSD content model:
 *   prstTxWarp?, EG_TextAutofit?, scene3d?, EG_Text3D?
 * followed by 19 attributes.
 */
export interface BodyPropertiesOptions {
  rotation?: number;
  spcFirstLastPara?: boolean;
  vertOverflow?: (typeof TextVertOverflowType)[keyof typeof TextVertOverflowType];
  horzOverflow?: (typeof TextHorzOverflowType)[keyof typeof TextHorzOverflowType];
  vert?: (typeof TextVerticalType)[keyof typeof TextVerticalType];
  wrap?: (typeof TextBodyWrappingType)[keyof typeof TextBodyWrappingType];
  lIns?: number | UniversalMeasure;
  tIns?: number | UniversalMeasure;
  rIns?: number | UniversalMeasure;
  bIns?: number | UniversalMeasure;
  numCol?: number;
  spcCol?: number | UniversalMeasure;
  rtlCol?: boolean;
  fromWordArt?: boolean;
  anchor?: (typeof VerticalAnchor)[keyof typeof VerticalAnchor];
  anchorCtr?: boolean;
  forceAA?: boolean;
  upright?: boolean;
  compatLnSpc?: boolean;

  /** Vertical anchor position (alias for `anchor`). */
  verticalAnchor?: VerticalAnchor;
  /** Margins shorthand. */
  margins?: {
    top?: number | UniversalMeasure;
    bottom?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
    right?: number | UniversalMeasure;
  };

  prstTxWarp?: PresetTextShapeOptions;
  /** Disable autofit (EG_TextAutofit choice). */
  noAutoFit?: boolean;
  /** Normal autofit (EG_TextAutofit choice). */
  normAutofit?: NormalAutofitOptions;
  /** Shape autofit (EG_TextAutofit choice). */
  spAutoFit?: boolean;
  scene3d?: Scene3DOptions;
  sp3d?: Shape3DOptions;
  flatTx?: FlatTextOptions;
}

// ── Helpers ──

const filterAttrs = (
  options: Record<string, unknown>,
): Record<string, string | number | boolean> | undefined => {
  const attrs: Record<string, string | number | boolean> = {};
  for (const [key, value] of Object.entries(options)) {
    if (value !== undefined) {
      attrs[key] = value as string | number | boolean;
    }
  }
  return Object.keys(attrs).length > 0 ? attrs : undefined;
};

const createPresetTextShape = (options: PresetTextShapeOptions): string => {
  const adjChildren: string[] = [];
  if (options.adjustments) {
    for (const adj of options.adjustments) {
      adjChildren.push(`<a:gd name="${adj.name}" fmla="${adj.formula}"/>`);
    }
  }
  return element(
    "a:prstTxWarp",
    { prst: options.preset },
    adjChildren.length > 0 ? [element("a:avLst", undefined, adjChildren)] : undefined,
  );
};

// ── Factory ──

/**
 * Create a text body properties element. tagName defaults to a:bodyPr (PPTX/XLSX);
 * DOCX passes "wps:bodyPr".
 */
export const createBodyProperties = (
  options: BodyPropertiesOptions = {},
  tagName = "a:bodyPr",
): string => {
  const anchor = options.anchor ?? options.verticalAnchor;
  const lIns = options.lIns ?? options.margins?.left;
  const tIns = options.tIns ?? options.margins?.top;
  const rIns = options.rIns ?? options.margins?.right;
  const bIns = options.bIns ?? options.margins?.bottom;

  const attrs = filterAttrs({
    rot: options.rotation,
    spcFirstLastPara: options.spcFirstLastPara,
    vertOverflow: options.vertOverflow,
    horzOverflow: options.horzOverflow,
    vert: options.vert,
    wrap: options.wrap,
    lIns: lIns !== undefined ? convertToEmu(lIns) : undefined,
    tIns: tIns !== undefined ? convertToEmu(tIns) : undefined,
    rIns: rIns !== undefined ? convertToEmu(rIns) : undefined,
    bIns: bIns !== undefined ? convertToEmu(bIns) : undefined,
    numCol: options.numCol,
    spcCol: options.spcCol !== undefined ? convertToEmu(options.spcCol) : undefined,
    rtlCol: options.rtlCol,
    fromWordArt: options.fromWordArt,
    anchor,
    anchorCtr: options.anchorCtr,
    forceAA: options.forceAA,
    upright: options.upright,
    compatLnSpc: options.compatLnSpc,
  });

  const children: string[] = [];
  if (options.prstTxWarp) children.push(createPresetTextShape(options.prstTxWarp));

  // EG_TextAutofit (mutually exclusive)
  if (options.noAutoFit) {
    children.push(`<a:noAutofit/>`);
  } else if (options.normAutofit) {
    const normAttrs = filterAttrs({
      fontScale: options.normAutofit.fontScale,
      lnSpcReduction: options.normAutofit.lnSpcReduction,
    });
    children.push(element("a:normAutofit", normAttrs));
  } else if (options.spAutoFit) {
    children.push(`<a:spAutoFit/>`);
  }

  if (options.scene3d) children.push(createScene3D(options.scene3d));

  // EG_Text3D (mutually exclusive)
  if (options.sp3d) {
    children.push(createShape3D(options.sp3d));
  } else if (options.flatTx) {
    const flatAttrs = filterAttrs({ z: options.flatTx.z });
    children.push(element("a:flatTx", flatAttrs));
  }

  return element(tagName, attrs, children.length > 0 ? children : undefined);
};

// ── Parse ──

/** Parse a bodyPr element into BodyPropertiesOptions. */
export const parseBodyProperties = (el: Element, ctx: ReadContext): BodyPropertiesOptions => {
  const result: BodyPropertiesOptions = {};

  const rotation = attrNum(el, "rot");
  if (rotation !== undefined) result.rotation = rotation;
  const spcFirstLastPara = attrBool(el, "spcFirstLastPara");
  if (spcFirstLastPara !== undefined) result.spcFirstLastPara = spcFirstLastPara;
  const vertOverflow = attr(el, "vertOverflow");
  if (vertOverflow !== undefined)
    result.vertOverflow = vertOverflow as BodyPropertiesOptions["vertOverflow"];
  const horzOverflow = attr(el, "horzOverflow");
  if (horzOverflow !== undefined)
    result.horzOverflow = horzOverflow as BodyPropertiesOptions["horzOverflow"];
  const vert = attr(el, "vert");
  if (vert !== undefined) result.vert = vert as BodyPropertiesOptions["vert"];
  const wrap = attr(el, "wrap");
  if (wrap !== undefined) result.wrap = wrap as BodyPropertiesOptions["wrap"];
  const lIns = attrMeasure(el, "lIns");
  if (lIns !== undefined) result.lIns = lIns as number | UniversalMeasure;
  const tIns = attrMeasure(el, "tIns");
  if (tIns !== undefined) result.tIns = tIns as number | UniversalMeasure;
  const rIns = attrMeasure(el, "rIns");
  if (rIns !== undefined) result.rIns = rIns as number | UniversalMeasure;
  const bIns = attrMeasure(el, "bIns");
  if (bIns !== undefined) result.bIns = bIns as number | UniversalMeasure;
  const numCol = attrNum(el, "numCol");
  if (numCol !== undefined) result.numCol = numCol;
  const spcCol = attrNum(el, "spcCol");
  if (spcCol !== undefined) result.spcCol = spcCol;
  const rtlCol = attrBool(el, "rtlCol");
  if (rtlCol !== undefined) result.rtlCol = rtlCol;
  const fromWordArt = attrBool(el, "fromWordArt");
  if (fromWordArt !== undefined) result.fromWordArt = fromWordArt;
  const anchor = attr(el, "anchor");
  if (anchor !== undefined) result.anchor = anchor as BodyPropertiesOptions["anchor"];
  const anchorCtr = attrBool(el, "anchorCtr");
  if (anchorCtr !== undefined) result.anchorCtr = anchorCtr;
  const forceAA = attrBool(el, "forceAA");
  if (forceAA !== undefined) result.forceAA = forceAA;
  const upright = attrBool(el, "upright");
  if (upright !== undefined) result.upright = upright;
  const compatLnSpc = attrBool(el, "compatLnSpc");
  if (compatLnSpc !== undefined) result.compatLnSpc = compatLnSpc;

  // EG_TextAutofit (mutually exclusive)
  if (findChild(el, "a:noAutofit")) {
    result.noAutoFit = true;
  } else {
    const norm = findChild(el, "a:normAutofit");
    if (norm) {
      const normOpts: NormalAutofitOptions = {};
      const fontScale = attrNum(norm, "fontScale");
      if (fontScale !== undefined) normOpts.fontScale = fontScale;
      const lnSpcReduction = attrNum(norm, "lnSpcReduction");
      if (lnSpcReduction !== undefined) normOpts.lnSpcReduction = lnSpcReduction;
      result.normAutofit = normOpts;
    } else if (findChild(el, "a:spAutoFit")) {
      result.spAutoFit = true;
    }
  }

  // a:prstTxWarp (CT_PresetTextShape)
  const prstTxWarp = findChild(el, "a:prstTxWarp");
  if (prstTxWarp) {
    const preset = attr(prstTxWarp, "prst") ?? "";
    const avLst = findChild(prstTxWarp, "a:avLst");
    const adjustments: { name: string; formula: string }[] = [];
    for (const gd of avLst?.elements ?? []) {
      if (gd.type === "element" && gd.name === "a:gd") {
        adjustments.push({ name: attr(gd, "name") ?? "", formula: attr(gd, "fmla") ?? "" });
      }
    }
    result.prstTxWarp = { preset, ...(adjustments.length > 0 ? { adjustments } : {}) };
  }

  // a:scene3d
  const scene3d = findChild(el, "a:scene3d");
  if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, ctx);

  // EG_Text3D (mutually exclusive: sp3d | flatTx)
  const sp3d = findChild(el, "a:sp3d");
  if (sp3d) {
    result.sp3d = shape3DDesc.parse(sp3d, ctx);
  } else {
    const flatTx = findChild(el, "a:flatTx");
    if (flatTx) {
      const z = attrNum(flatTx, "z");
      result.flatTx = z !== undefined ? { z } : {};
    }
  }

  return result as BodyPropertiesOptions;
};

// ── Descriptor (a:bodyPr instance) ──

export const bodyPropertiesDesc: CustomDescriptor<
  BodyPropertiesOptions,
  WriteContext,
  BodyPropertiesOptions
> = {
  kind: "custom",
  stringify(opts) {
    return createBodyProperties(opts);
  },
  parse(el, ctx) {
    return parseBodyProperties(el, ctx);
  },
};
