/**
 * Transform 2D descriptor for DrawingML shapes.
 *
 * @module
 */

import { attrMeasure, findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import { convertToEmu } from "../util/converters";
import type { GroupTransform2DOptions, Transform2DOptions } from "./transform";

// ── Transform2D descriptor ──
// x, y, width, height: number (EMU) | string (UniversalMeasure)

export const transform2DDesc: CustomDescriptor<Transform2DOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const parts: string[] = [];

    // Attributes
    const attrParts: string[] = [];
    if (opts.flipHorizontal !== undefined) attrParts.push(`flipH="${opts.flipHorizontal ? 1 : 0}"`);
    if (opts.flipVertical !== undefined) attrParts.push(`flipV="${opts.flipVertical ? 1 : 0}"`);
    if (opts.rotation !== undefined) attrParts.push(`rot="${opts.rotation}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    // Child elements: off, ext
    if (opts.x !== undefined || opts.y !== undefined) {
      const x = opts.x !== undefined ? convertToEmu(opts.x) : 0;
      const y = opts.y !== undefined ? convertToEmu(opts.y) : 0;
      parts.push(`<a:off x="${x}" y="${y}"/>`);
    }
    if (opts.width !== undefined || opts.height !== undefined) {
      const cx = opts.width !== undefined ? convertToEmu(opts.width) : 0;
      const cy = opts.height !== undefined ? convertToEmu(opts.height) : 0;
      parts.push(`<a:ext cx="${cx}" cy="${cy}"/>`);
    }

    if (parts.length === 0 && !attrStr) return undefined;
    if (parts.length === 0) return `<a:xfrm${attrStr}/>`;
    return `<a:xfrm${attrStr}>${parts.join("")}</a:xfrm>`;
  },
  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Attributes
    if (el.attributes) {
      if (el.attributes["flipH"] !== undefined)
        result.flipHorizontal = el.attributes["flipH"] === 1 || el.attributes["flipH"] === "1";
      if (el.attributes["flipV"] !== undefined)
        result.flipVertical = el.attributes["flipV"] === 1 || el.attributes["flipV"] === "1";
      if (el.attributes["rot"] !== undefined) result.rotation = Number(el.attributes["rot"]);
    }

    // Child: off — a:off x/y are ST_Coordinate (number EMU | UniversalMeasure)
    const off = findChild(el, "a:off");
    if (off?.attributes) {
      const xv = attrMeasure(off, "x");
      result.x = xv ?? 0;
      const yv = attrMeasure(off, "y");
      result.y = yv ?? 0;
    }

    // Child: ext
    const ext = findChild(el, "a:ext");
    if (ext?.attributes) {
      const cxv = attrMeasure(ext, "cx");
      result.width = cxv ?? 0;
      const cyv = attrMeasure(ext, "cy");
      result.height = cyv ?? 0;
    }

    return result as Transform2DOptions;
  },
};

// ── GroupTransform2D descriptor ──

export const groupTransform2DDesc: CustomDescriptor<GroupTransform2DOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const base = transform2DDesc.stringify(opts, ctx) ?? "<a:xfrm/>";

    // Append chOff and chExt children before closing tag (values are already in EMU)
    const chOff = `<a:chOff x="${opts.childOffsetX ?? 0}" y="${opts.childOffsetY ?? 0}"/>`;
    const chExt = `<a:chExt cx="${opts.childExtentWidth ?? 0}" cy="${opts.childExtentHeight ?? 0}"/>`;

    if (base.endsWith("/>")) {
      return `<a:xfrm>${chOff}${chExt}</a:xfrm>`;
    }
    return base.replace(/<\/a:xfrm>$/, `${chOff}${chExt}</a:xfrm>`);
  },
  parse(el, ctx) {
    const result = transform2DDesc.parse(el, ctx) as Record<string, unknown>;

    const chOff = findChild(el, "a:chOff");
    if (chOff?.attributes) {
      const cx = attrMeasure(chOff, "x");
      result.childOffsetX = cx ?? 0;
      const cy = attrMeasure(chOff, "y");
      result.childOffsetY = cy ?? 0;
    }

    const chExt = findChild(el, "a:chExt");
    if (chExt?.attributes) {
      const ccx = attrMeasure(chExt, "cx");
      result.childExtentWidth = ccx ?? 0;
      const ccy = attrMeasure(chExt, "cy");
      result.childExtentHeight = ccy ?? 0;
    }

    return result as GroupTransform2DOptions;
  },
};
