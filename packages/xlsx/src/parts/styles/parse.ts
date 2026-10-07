/**
 * Styles — parse helpers for xl/styles.xml sub-elements.
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { parseColorOptions } from "../../shared/color";
import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellProtectionOptions,
  CellFillOptions,
  FontOptions,
  CellGradientStopOptions,
} from "./types";

export function parseFont(el: XmlElement): FontOptions {
  const result: Partial<FontOptions> = {};
  const childOrder: string[] = [];
  // CT_BooleanProperty children (b/i/strike/…) default val to true, so an
  // explicit val="0" is a real "off" statement — keep it as false rather than
  // collapsing to the attribute-absent default.
  const boolProp = (prop: XmlElement): boolean | undefined => parseOnOff(attr(prop, "val")) ?? true;
  for (const child of el.elements ?? []) {
    if (child.name !== undefined) childOrder.push(child.name);
    switch (child.name) {
      case "b":
        result.bold = boolProp(child);
        result.boldRaw = attr(child, "val");
        break;
      case "i":
        result.italic = boolProp(child);
        result.italicRaw = attr(child, "val");
        break;
      case "u":
        // ST_UnderlineValues: "none" is the only off spelling; every other
        // value (or the omitted default) means some underline is on.
        const underline = attr(child, "val");
        switch (underline) {
          case undefined:
            result.underline = true;
            break;
          case "none":
            result.underline = false;
            break;
          case "single":
          case "double":
          case "singleAccounting":
          case "doubleAccounting":
            result.underline = underline;
            break;
        }
        break;
      case "strike":
        result.strike = boolProp(child);
        result.strikeRaw = attr(child, "val");
        break;
      case "outline":
        result.outline = boolProp(child);
        result.outlineRaw = attr(child, "val");
        break;
      case "shadow":
        result.shadow = boolProp(child);
        result.shadowRaw = attr(child, "val");
        break;
      case "condense":
        result.condense = boolProp(child);
        result.condenseRaw = attr(child, "val");
        break;
      case "extend":
        result.extend = boolProp(child);
        result.extendRaw = attr(child, "val");
        break;
      case "sz":
        result.size = attrNum(child, "val");
        result.sizeRaw = attr(child, "val");
        break;
      case "color":
        result.color = parseColorOptions(child);
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
  if (childOrder.length > 1) result.childOrder = childOrder;
  return result as FontOptions;
}

export function parseFill(el: XmlElement): CellFillOptions {
  const legacyPattern = findChild(el, "pattern");
  if (legacyPattern) {
    const result: CellFillOptions = { type: "pattern", legacyPatternElement: true };
    const patternType = attr(legacyPattern, "patternType");
    if (patternType) {
      result.legacyPatternType = patternType;
      if (["none", "solid", "gray125"].includes(patternType))
        result.patternType = patternType as CellFillOptions["patternType"];
    }
    const fg = findChild(legacyPattern, "fgColor");
    if (fg) {
      result.foregroundColor = parseColorOptions(fg);
    }
    const bg = findChild(legacyPattern, "bgColor");
    if (bg) {
      result.backgroundColor = parseColorOptions(bg);
    }
    return result;
  }

  const patternFill = findChild(el, "patternFill");
  if (patternFill) {
    const result: CellFillOptions = {};
    const patternType = attr(patternFill, "patternType");
    result.patternTypeDeclared = patternType !== undefined;
    if (patternType) result.patternType = patternType as CellFillOptions["patternType"];
    const fg = findChild(patternFill, "fgColor");
    if (fg) {
      result.foregroundColor = parseColorOptions(fg);
    }
    const bg = findChild(patternFill, "bgColor");
    if (bg) {
      result.backgroundColor = parseColorOptions(bg);
    }
    return result;
  }

  const gradientFill = findChild(el, "gradientFill");
  if (gradientFill) {
    const result: CellFillOptions = { type: "gradient" };
    const gType = attr(gradientFill, "type");
    if (gType) result.gradientType = gType as CellFillOptions["gradientType"];
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
    const stops: CellGradientStopOptions[] = [];
    for (const s of gradientFill.elements ?? []) {
      if (s.name !== "stop") continue;
      const pos = attrNum(s, "position");
      const color = findChild(s, "color");
      if (pos === undefined || !color) continue;
      const stop: CellGradientStopOptions = { position: pos };
      stop.color = parseColorOptions(color);
      stops.push(stop);
    }
    if (stops.length > 0) result.stops = stops;
    return result;
  }

  return {};
}

export function parseBorder(el: XmlElement): BorderSideOptions {
  const result: BorderSideOptions = {};
  const childOrder: string[] = [];
  if (attr(el, "diagonalUp") !== undefined)
    result.diagonalUp = parseOnOff(attr(el, "diagonalUp")) ?? true;
  if (attr(el, "diagonalDown") !== undefined)
    result.diagonalDown = parseOnOff(attr(el, "diagonalDown")) ?? true;
  if (attr(el, "outline") !== undefined) result.outline = parseOnOff(attr(el, "outline")) ?? true;

  for (const sideEl of el.elements ?? []) {
    const side = sideEl.name;
    if (
      side !== undefined &&
      (
        [
          "left",
          "right",
          "top",
          "bottom",
          "diagonal",
          "start",
          "end",
          "vertical",
          "horizontal",
        ] as const
      ).includes(side as "left")
    ) {
      childOrder.push(side);
      // Presence-preserving: an empty <left/> stays as left: {} so stringify
      // re-emits it; Excel always writes the five cell sides, dxf adds
      // vertical/horizontal — both round-trip byte-identically.
      const opts: BorderOptions = {};
      const style = attr(sideEl, "style");
      if (style !== undefined) opts.style = style as BorderOptions["style"];
      const color = findChild(sideEl, "color");
      if (color) {
        opts.color = parseColorOptions(color);
      }
      result[
        side as
          | "left"
          | "right"
          | "top"
          | "bottom"
          | "diagonal"
          | "start"
          | "end"
          | "vertical"
          | "horizontal"
      ] = opts;
    }
  }

  if (childOrder.length > 1) result.childOrder = childOrder;

  return result;
}

export function parseAlignment(el: XmlElement): AlignmentOptions {
  const result: AlignmentOptions = {};
  const h = attr(el, "horizontal");
  if (h) result.horizontal = h as AlignmentOptions["horizontal"];
  const v = attr(el, "vertical");
  if (v) result.vertical = v as AlignmentOptions["vertical"];
  const wrapText = parseOnOff(attr(el, "wrapText"));
  if (wrapText !== undefined) result.wrapText = wrapText;
  const rotation = attrNum(el, "textRotation");
  if (rotation !== undefined) result.textRotation = rotation;
  const indent = attrNum(el, "indent");
  if (indent !== undefined) result.indent = indent;
  const indentRaw = attr(el, "indent");
  if (indentRaw !== undefined && indent === undefined) result.indentRaw = indentRaw;
  const relativeIndent = attrNum(el, "relativeIndent");
  if (relativeIndent !== undefined) result.relativeIndent = relativeIndent;
  const justifyLastLine = parseOnOff(attr(el, "justifyLastLine"));
  if (justifyLastLine !== undefined) result.justifyLastLine = justifyLastLine;
  const shrinkToFit = parseOnOff(attr(el, "shrinkToFit"));
  if (shrinkToFit !== undefined) result.shrinkToFit = shrinkToFit;
  const readingOrder = attrNum(el, "readingOrder");
  if (readingOrder !== undefined) result.readingOrder = readingOrder;
  const mergeCell = parseOnOff(attr(el, "mergeCell"));
  if (mergeCell !== undefined) result.mergeCell = mergeCell;
  return result;
}

export function parseProtection(el: XmlElement): CellProtectionOptions {
  const result: Partial<CellProtectionOptions> = {};
  const locked = attr(el, "locked");
  if (locked !== undefined) result.locked = parseOnOff(locked) ?? true;
  const hidden = attr(el, "hidden");
  if (hidden !== undefined) result.hidden = parseOnOff(hidden) ?? true;
  return result as CellProtectionOptions;
}

/**
 * Read a CT_RgbColor `@rgb`, stripping the alpha prefix when present
 * (FF000000 → 000000). Used by legacy color-list extensions, not CT_Color.
 */
export function parseColorHex(el: XmlElement): string | undefined {
  const rgb = attr(el, "rgb");
  if (rgb) {
    // Strip alpha prefix if present (FF000000 → 000000)
    return rgb.length === 8 ? rgb.slice(2) : rgb;
  }
  return undefined;
}
