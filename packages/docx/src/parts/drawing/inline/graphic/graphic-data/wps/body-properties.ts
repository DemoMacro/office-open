/**
 * Text body properties for DrawingML shapes.
 *
 * Provides CT_TextBodyProperties — defines how text is laid out within a shape,
 * including margins, alignment, autofit, vertical text, overflow, columns, and 3D.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TextBodyProperties
 *
 * @module
 */
import type { UniversalMeasure } from "@office-open/core";
import { convertToEmu } from "@office-open/core";
import { attr, attrBool, attrNum, element, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { Scene3DOptions } from "../pic/three-d/scene-3d";
import { createScene3D } from "../pic/three-d/scene-3d";
import type { Shape3DOptions } from "../pic/three-d/shape-3d";
import { createShape3D } from "../pic/three-d/shape-3d";

// ─── Constants ──────────────────────────────────────────────────────────────

/**
 * Text anchoring type (ST_TextAnchoringType).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TextAnchoringType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="t"/>
 *     <xsd:enumeration value="ctr"/>
 *     <xsd:enumeration value="b"/>
 *     <xsd:enumeration value="just"/>
 *     <xsd:enumeration value="dist"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export enum VerticalAnchor {
  TOP = "t",
  CENTER = "ctr",
  BOTTOM = "b",
  JUSTIFY = "just",
  DISTRIBUTED = "dist",
}

/**
 * Text vertical overflow type (ST_TextVertOverflowType).
 */
export const TextVertOverflowType = {
  OVERFLOW: "overflow",
  ELLIPSIS: "ellipsis",
  CLIP: "clip",
} as const;

/**
 * Text horizontal overflow type (ST_TextHorzOverflowType).
 */
export const TextHorzOverflowType = {
  OVERFLOW: "overflow",
  CLIP: "clip",
} as const;

/**
 * Text vertical type (ST_TextVerticalType).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TextVerticalType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="horz"/>
 *     <xsd:enumeration value="vert"/>
 *     <xsd:enumeration value="vert270"/>
 *     <xsd:enumeration value="wordArtVert"/>
 *     <xsd:enumeration value="eaVert"/>
 *     <xsd:enumeration value="mongolianVert"/>
 *     <xsd:enumeration value="wordArtVertRtl"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const TextVerticalType = {
  HORIZONTAL: "horz",
  VERTICAL: "vert",
  VERTICAL_270: "vert270",
  WORD_ART_VERTICAL: "wordArtVert",
  EAST_ASIAN_VERTICAL: "eaVert",
  MONGOLIAN_VERTICAL: "mongolianVert",
  WORD_ART_VERTICAL_RTL: "wordArtVertRtl",
} as const;

/**
 * Text body wrapping type (ST_TextWrappingType).
 *
 * This is different from text wrapping around shapes (ST_WrapText).
 */
export const TextBodyWrappingType = {
  NONE: "none",
  SQUARE: "square",
} as const;

// ─── Sub-element Options ────────────────────────────────────────────────────

/**
 * Normal autofit options (CT_TextNormalAutofit).
 *
 * Scales text and line spacing to fit the shape.
 */
export interface NormalAutofitOptions {
  /** Font scale percentage (e.g., 100000 = 100%) */
  fontScale?: number;
  /** Line spacing reduction percentage */
  lnSpcReduction?: number;
}

/**
 * Preset text warp options (CT_PresetTextShape).
 */
export interface PresetTextShapeOptions {
  /** Preset shape type (e.g., "textArchUp", "textCircle") */
  preset: string;
  /** Optional adjustment values */
  adjustments?: { name: string; formula: string }[];
}

/**
 * Flat text options (CT_FlatText).
 */
export interface FlatTextOptions {
  /** Z-offset in EMUs */
  z?: number;
}

// ─── Main Options ───────────────────────────────────────────────────────────

/**
 * Options for text body properties (CT_TextBodyProperties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TextBodyProperties">
 *   <xsd:sequence>
 *     <xsd:element name="prstTxWarp" type="CT_PresetTextShape" minOccurs="0"/>
 *     <xsd:group ref="EG_TextAutofit" minOccurs="0"/>
 *     <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0"/>
 *     <xsd:group ref="EG_Text3D" minOccurs="0"/>
 *   </xsd:sequence>
 *   <!-- 19 attributes -->
 * </xsd:complexType>
 * ```
 */
export interface BodyPropertiesOptions {
  // ── Attributes ──

  /** Text rotation angle in 60,000ths of a degree */
  rotation?: number;
  /** Whether to use spcFirstLastPara behavior */
  spcFirstLastPara?: boolean;
  /** Vertical text overflow behavior */
  vertOverflow?: (typeof TextVertOverflowType)[keyof typeof TextVertOverflowType];
  /** Horizontal text overflow behavior */
  horzOverflow?: (typeof TextHorzOverflowType)[keyof typeof TextHorzOverflowType];
  /** Text vertical direction */
  vert?: (typeof TextVerticalType)[keyof typeof TextVerticalType];
  /** Text wrapping type */
  wrap?: (typeof TextBodyWrappingType)[keyof typeof TextBodyWrappingType];
  /** Left inset in EMUs or universal measure (e.g., "1cm", "0.5in") */
  lIns?: number | UniversalMeasure;
  /** Top inset in EMUs or universal measure */
  tIns?: number | UniversalMeasure;
  /** Right inset in EMUs or universal measure */
  rIns?: number | UniversalMeasure;
  /** Bottom inset in EMUs or universal measure */
  bIns?: number | UniversalMeasure;
  /** Number of text columns (1-16) */
  numCol?: number;
  /** Spacing between columns in EMUs or universal measure */
  spcCol?: number | UniversalMeasure;
  /** Whether columns are right-to-left */
  rtlCol?: boolean;
  /** Whether text is from WordArt */
  fromWordArt?: boolean;
  /** Text anchor position */
  anchor?: (typeof VerticalAnchor)[keyof typeof VerticalAnchor];
  /** Whether to anchor at center */
  anchorCtr?: boolean;
  /** Whether to force anti-aliasing */
  forceAA?: boolean;
  /** Whether text is upright (default false) */
  upright?: boolean;
  /** Whether to use compatible line spacing */
  compatLnSpc?: boolean;

  // ── Convenience aliases (backward compatible) ──

  /** Vertical anchor position (alias for `anchor`) */
  verticalAnchor?: VerticalAnchor;
  /** Margins shorthand */
  margins?: {
    top?: number | UniversalMeasure;
    bottom?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
    right?: number | UniversalMeasure;
  };

  // ── Child elements ──

  /** Preset text warp shape */
  prstTxWarp?: PresetTextShapeOptions;
  /** Disable autofit (EG_TextAutofit choice) */
  noAutoFit?: boolean;
  /** Normal autofit (EG_TextAutofit choice) */
  normAutofit?: NormalAutofitOptions;
  /** Shape autofit (EG_TextAutofit choice) */
  spAutoFit?: boolean;
  /** 3D scene */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (EG_Text3D choice) */
  sp3d?: Shape3DOptions;
  /** Flat text (EG_Text3D choice) */
  flatTx?: FlatTextOptions;
}

// ─── Internal Helpers ───────────────────────────────────────────────────────

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

// ─── Preset Text Warp ───────────────────────────────────────────────────────

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

// ─── Main Factory ───────────────────────────────────────────────────────────

/**
 * Creates a text body properties element (wps:bodyPr).
 *
 * Supports all 19 CT_TextBodyProperties attributes and all child elements
 * including preset text warp, autofit variants, 3D scene, and text 3D.
 *
 * @example
 * ```typescript
 * // Simple margins and anchor
 * createBodyProperties({
 *   margins: { top: 100, bottom: 100, left: 200, right: 200 },
 *   verticalAnchor: VerticalAnchor.CENTER,
 * });
 *
 * // Full options with vertical text and autofit
 * createBodyProperties({
 *   vert: TextVerticalType.VERTICAL,
 *   wrap: TextBodyWrappingType.NONE,
 *   normAutofit: { fontScale: 80000 },
 *   numCol: 2,
 *   anchor: VerticalAnchor.TOP,
 * });
 * ```
 */
export const createBodyProperties = (options: BodyPropertiesOptions = {}): string => {
  // Resolve anchor (direct `anchor` takes precedence over `verticalAnchor`)
  const anchor = options.anchor ?? options.verticalAnchor;

  // Resolve margins
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

  // Build children in XSD sequence order
  const children: string[] = [];

  // a:prstTxWarp
  if (options.prstTxWarp) {
    children.push(createPresetTextShape(options.prstTxWarp));
  }

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

  // a:scene3d
  if (options.scene3d) {
    children.push(createScene3D(options.scene3d));
  }

  // EG_Text3D (mutually exclusive)
  if (options.sp3d) {
    children.push(createShape3D(options.sp3d));
  } else if (options.flatTx) {
    const flatAttrs = filterAttrs({ z: options.flatTx.z });
    children.push(element("a:flatTx", flatAttrs));
  }

  return element("wps:bodyPr", attrs, children.length > 0 ? children : undefined);
};

// ─── Parse ──────────────────────────────────────────────────────────────────

/**
 * Parse a `wps:bodyPr` element into {@link BodyPropertiesOptions}.
 *
 * Reads the CT_TextBodyProperties attributes and the EG_TextAutofit child
 * (noAutofit/normAutofit/spAutoFit). prstTxWarp/scene3d/text-3D are not yet
 * parsed (later phase).
 */
export const parseBodyProperties = (el: Element): BodyPropertiesOptions => {
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
  const lIns = attrNum(el, "lIns");
  if (lIns !== undefined) result.lIns = lIns;
  const tIns = attrNum(el, "tIns");
  if (tIns !== undefined) result.tIns = tIns;
  const rIns = attrNum(el, "rIns");
  if (rIns !== undefined) result.rIns = rIns;
  const bIns = attrNum(el, "bIns");
  if (bIns !== undefined) result.bIns = bIns;
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

  return result as BodyPropertiesOptions;
};
