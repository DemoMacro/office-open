/**
 * Spreadsheet color shared by styles, conditional formats, and sheet
 * properties (CT_Color).
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import type { ArgbHexColor } from "@office-open/core";
import { attr, attrNum } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

/** Spreadsheet color (CT_Color). Exactly one channel — `rgb`, `theme` or `indexed` — is authoritative; `tint` qualifies `theme`. */
export interface ColorOptions {
  /** Automatic system color (CT_Color `@auto`) */
  auto?: boolean;
  /** Indexed legacy palette entry (CT_Color `@indexed`) */
  indexed?: number;
  /** ARGB hex with alpha, e.g. "FF4472C4" (CT_Color `@rgb`) */
  rgb?: ArgbHexColor;
  /** Theme palette index (CT_Color `@theme`) */
  theme?: number;
  /** Tint applied to the theme color, -1.0–1.0 (CT_Color `@tint`) */
  tint?: number;
}

/** Serialize a color's CT_Color attributes (without leading space). */
export function colorAttributes(color: ColorOptions): string {
  const parts: string[] = [];
  if (color.auto) parts.push('auto="1"');
  if (color.indexed !== undefined) parts.push(`indexed="${color.indexed}"`);
  if (color.rgb !== undefined) parts.push(`rgb="${color.rgb}"`);
  if (color.theme !== undefined) parts.push(`theme="${color.theme}"`);
  if (color.tint !== undefined) parts.push(`tint="${color.tint}"`);
  return parts.join(" ");
}

/** Parse CT_Color attributes into canonical options; undefined channels are omitted. */
export function parseColorOptions(el: XmlElement): ColorOptions {
  const color: ColorOptions = {};
  const legacyType = attr(el, "type");
  const legacyValue = attr(el, "val");
  if (legacyType === "rgb" && legacyValue !== undefined)
    color.rgb = String(legacyValue) as ArgbHexColor;
  else if (legacyType === "theme" && legacyValue !== undefined) {
    const theme = attrNum(el, "val");
    if (theme !== undefined) color.theme = theme;
  } else if (legacyType === "icv" && legacyValue !== undefined) {
    const indexed = attrNum(el, "val");
    if (indexed !== undefined) color.indexed = indexed;
  }
  const rgb = attr(el, "rgb");
  if (rgb !== undefined) color.rgb = String(rgb) as ArgbHexColor;
  const theme = attrNum(el, "theme");
  if (theme !== undefined) color.theme = theme;
  const tint = attrNum(el, "tint");
  if (tint !== undefined) color.tint = tint;
  const indexed = attrNum(el, "indexed");
  if (indexed !== undefined) color.indexed = indexed;
  if (parseOnOff(attr(el, "auto")) === true) color.auto = true;
  return color;
}
