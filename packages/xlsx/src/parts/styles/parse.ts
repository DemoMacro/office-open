/**
 * Styles — parse helpers for xl/styles.xml sub-elements.
 *
 * @module
 */
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellProtectionOptions,
  FillOptions,
  FontOptions,
  GradientStopOptions,
} from "./types";

export function parseFont(el: XmlElement): FontOptions {
  const result: Partial<FontOptions> = {};
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "b":
        result.bold = true;
        break;
      case "i":
        result.italic = true;
        break;
      case "u":
        result.underline = true;
        break;
      case "strike":
        result.strike = true;
        break;
      case "outline":
        result.outline = true;
        break;
      case "shadow":
        result.shadow = true;
        break;
      case "condense":
        result.condense = true;
        break;
      case "extend":
        result.extend = true;
        break;
      case "sz":
        result.size = attrNum(child, "val");
        break;
      case "color":
        result.color = parseColorHex(child);
        break;
      case "name":
        result.font = attr(child, "val") ?? undefined;
        break;
      case "charset":
        result.charset = attrNum(child, "val");
        break;
      case "family":
        result.family = attrNum(child, "val");
        break;
      case "vertAlign":
        result.vertAlign = (attr(child, "val") as FontOptions["vertAlign"]) ?? undefined;
        break;
      case "scheme":
        result.scheme = (attr(child, "val") as FontOptions["scheme"]) ?? undefined;
        break;
    }
  }
  return result as FontOptions;
}

export function parseFill(el: XmlElement): FillOptions {
  const patternFill = findChild(el, "patternFill");
  if (patternFill) {
    const result: FillOptions = {};
    const patternType = attr(patternFill, "patternType");
    if (patternType) result.patternType = patternType;
    const fg = findChild(patternFill, "fgColor");
    if (fg) result.color = parseColorHex(fg);
    const bg = findChild(patternFill, "bgColor");
    if (bg) result.bgColor = parseColorHex(bg);
    const indexed = fg ? attrNum(fg, "indexed") : undefined;
    if (indexed !== undefined) result.colorIndexed = indexed;
    return result;
  }

  const gradientFill = findChild(el, "gradientFill");
  if (gradientFill) {
    const result: FillOptions = { type: "gradient" };
    const gType = attr(gradientFill, "type");
    if (gType) result.gradientType = gType as FillOptions["gradientType"];
    const degree = attrNum(gradientFill, "degree");
    if (degree !== undefined) result.gradientDegree = degree;
    const left = attrNum(gradientFill, "left");
    if (left !== undefined) result.gradientLeft = left;
    const right = attrNum(gradientFill, "right");
    if (right !== undefined) result.gradientRight = right;
    const top = attrNum(gradientFill, "top");
    if (top !== undefined) result.gradientTop = top;
    const bottom = attrNum(gradientFill, "bottom");
    if (bottom !== undefined) result.gradientBottom = bottom;
    const stops: GradientStopOptions[] = [];
    for (const s of gradientFill.elements ?? []) {
      if (s.name !== "stop") continue;
      const pos = attrNum(s, "position");
      const color = findChild(s, "color");
      if (pos !== undefined && color) {
        stops.push({ position: pos, color: parseColorHex(color) ?? "" });
      }
    }
    if (stops.length > 0) result.stops = stops;
    return result;
  }

  return {};
}

export function parseBorder(el: XmlElement): BorderSideOptions {
  const result: BorderSideOptions = {};
  if (String(attr(el, "diagonalUp")) === "1") result.diagonalUp = true;
  if (String(attr(el, "diagonalDown")) === "1") result.diagonalDown = true;

  for (const side of [
    "left",
    "right",
    "top",
    "bottom",
    "diagonal",
    "start",
    "end",
    "vertical",
    "horizontal",
  ] as const) {
    const sideEl = findChild(el, side);
    if (sideEl) {
      const opts: BorderOptions = {};
      const style = attr(sideEl, "style");
      if (style) opts.style = style as BorderOptions["style"];
      const color = findChild(sideEl, "color");
      if (color) opts.color = parseColorHex(color);
      if (Object.keys(opts).length > 0) result[side] = opts;
    }
  }

  return result;
}

export function parseAlignment(el: XmlElement): AlignmentOptions {
  const result: AlignmentOptions = {};
  const h = attr(el, "horizontal");
  if (h) result.horizontal = h as AlignmentOptions["horizontal"];
  const v = attr(el, "vertical");
  if (v) result.vertical = v as AlignmentOptions["vertical"];
  if (String(attr(el, "wrapText")) === "1") result.wrapText = true;
  const rotation = attrNum(el, "textRotation");
  if (rotation !== undefined) result.textRotation = rotation;
  const indent = attrNum(el, "indent");
  if (indent !== undefined) result.indent = indent;
  const relativeIndent = attrNum(el, "relativeIndent");
  if (relativeIndent !== undefined) result.relativeIndent = relativeIndent;
  if (String(attr(el, "justifyLastLine")) === "1") result.justifyLastLine = true;
  if (String(attr(el, "shrinkToFit")) === "1") result.shrinkToFit = true;
  const readingOrder = attrNum(el, "readingOrder");
  if (readingOrder !== undefined) result.readingOrder = readingOrder;
  return result;
}

export function parseProtection(el: XmlElement): CellProtectionOptions {
  const result: Partial<CellProtectionOptions> = {};
  const locked = attr(el, "locked");
  if (locked !== undefined) result.locked = locked !== "0";
  const hidden = attr(el, "hidden");
  if (hidden !== undefined) result.hidden = hidden !== "0";
  return result as CellProtectionOptions;
}

function parseColorHex(el: XmlElement): string | undefined {
  const rgb = attr(el, "rgb");
  if (rgb) {
    // Strip alpha prefix if present (FF000000 → 000000)
    return rgb.length === 8 ? rgb.slice(2) : rgb;
  }
  return undefined;
}
