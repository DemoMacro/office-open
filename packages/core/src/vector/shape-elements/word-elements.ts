/**
 * WordprocessingML drawing w10: elements — the EG_ShapeElements members from
 * vml-wordprocessingDrawing.xsd: wrap, anchorlock and the four borders.
 *
 * Reference: ISO/IEC 29500-4, vml-wordprocessingDrawing.xsd.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  type VmlAttrSpec,
  type VmlTrueFalseBlank,
} from "../attributes";

/** ST_WrapType. */
export type VmlWrapType = "topAndBottom" | "square" | "none" | "tight" | "through";

/** ST_WrapSide. */
export type VmlWrapSide = "both" | "left" | "right" | "largest";

/** ST_HorizontalAnchor. */
export type VmlHorizontalAnchor = "margin" | "page" | "text" | "char";

/** ST_VerticalAnchor. */
export type VmlVerticalAnchor = "margin" | "page" | "text" | "line";

/** ST_BorderType — the 27 legacy border presets. */
export type VmlBorderType =
  | "none"
  | "single"
  | "thick"
  | "double"
  | "hairline"
  | "dot"
  | "dash"
  | "dotDash"
  | "dashDotDot"
  | "triple"
  | "thinThickSmall"
  | "thickThinSmall"
  | "thickBetweenThinSmall"
  | "thinThick"
  | "thickThin"
  | "thickBetweenThin"
  | "thinThickLarge"
  | "thickThinLarge"
  | "thickBetweenThinLarge"
  | "wave"
  | "doubleWave"
  | "dashedSmall"
  | "dashDotStroked"
  | "threeDEmboss"
  | "threeDEngrave"
  | "HTMLOutset"
  | "HTMLInset";

/** w10:wrap options (CT_Wrap). */
export interface VmlWrapOptions {
  type?: VmlWrapType;
  side?: VmlWrapSide;
  anchorx?: VmlHorizontalAnchor;
  anchory?: VmlVerticalAnchor;
}

const WRAP_ATTRS: readonly VmlAttrSpec[] = [
  { field: "type", attr: "type", kind: "string" },
  { field: "side", attr: "side", kind: "string" },
  { field: "anchorx", attr: "anchorx", kind: "string" },
  { field: "anchory", attr: "anchory", kind: "string" },
];

/** Serialize w10:wrap. */
export function stringifyVmlWrap(opts: VmlWrapOptions): string {
  return `<w10:wrap${stringifyVmlAttributes(opts as Record<string, unknown>, WRAP_ATTRS)}/>`;
}

/** Parse a w10:wrap element. */
export function parseVmlWrap(el: XmlElement): VmlWrapOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, WRAP_ATTRS, out);
  return out as VmlWrapOptions;
}

/** w10:anchorlock options (CT_AnchorLock) — empty marker element. */
export interface VmlAnchorLockOptions {}

/** Serialize w10:anchorlock. */
export function stringifyVmlAnchorLock(_opts: VmlAnchorLockOptions): string {
  return "<w10:anchorlock/>";
}

/** Parse a w10:anchorlock element. */
export function parseVmlAnchorLock(_el: XmlElement): VmlAnchorLockOptions {
  return {};
}

/** w10:bordertop/borderleft/borderright/borderbottom options (CT_Border). */
export interface VmlBorderOptions {
  type?: VmlBorderType;
  /** Border width in 1/8 pt units. */
  width?: number;
  shadow?: VmlTrueFalseBlank;
}

const BORDER_ATTRS: readonly VmlAttrSpec[] = [
  { field: "type", attr: "type", kind: "string" },
  { field: "width", attr: "width", kind: "number" },
  { field: "shadow", attr: "shadow", kind: "trueFalseBlank" },
];

/** The four w10: border element names. */
const VML_BORDER_TAGS = [
  "w10:bordertop",
  "w10:borderleft",
  "w10:borderright",
  "w10:borderbottom",
] as const;

/** Serialize a w10: border element. */
export function stringifyVmlBorder(
  tag: (typeof VML_BORDER_TAGS)[number],
  opts: VmlBorderOptions,
): string {
  return `<${tag}${stringifyVmlAttributes(opts as Record<string, unknown>, BORDER_ATTRS)}/>`;
}

/** Parse a w10: border element. */
export function parseVmlBorder(el: XmlElement): VmlBorderOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, BORDER_ATTRS, out);
  return out as VmlBorderOptions;
}
