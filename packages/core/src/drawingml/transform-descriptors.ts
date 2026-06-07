/**
 * Transform 2D descriptor for DrawingML shapes.
 *
 * @module
 */

import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import type { GroupTransform2DOptions, Transform2DOptions } from "./transform";

// ── Transform2D descriptor ──

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
      parts.push(`<a:off x="${opts.x ?? 0}" y="${opts.y ?? 0}"/>`);
    }
    if (opts.width !== undefined || opts.height !== undefined) {
      parts.push(`<a:ext cx="${opts.width ?? 0}" cy="${opts.height ?? 0}"/>`);
    }

    if (parts.length === 0 && !attrStr) return undefined;
    if (parts.length === 0) return `<a:xfrm${attrStr}/>`;
    return `<a:xfrm${attrStr}>${parts.join("")}</a:xfrm>`;
  },
  parse(el, _ctx) {
    const result: Partial<Transform2DOptions> = {};

    // Attributes
    if (el.attributes) {
      if (el.attributes["flipH"] !== undefined)
        result.flipHorizontal = el.attributes["flipH"] === 1 || el.attributes["flipH"] === "1";
      if (el.attributes["flipV"] !== undefined)
        result.flipVertical = el.attributes["flipV"] === 1 || el.attributes["flipV"] === "1";
      if (el.attributes["rot"] !== undefined) result.rotation = Number(el.attributes["rot"]);
    }

    // Child: off
    const off = findChild(el, "a:off");
    if (off?.attributes) {
      result.x = Number(off.attributes["x"] ?? 0);
      result.y = Number(off.attributes["y"] ?? 0);
    }

    // Child: ext
    const ext = findChild(el, "a:ext");
    if (ext?.attributes) {
      result.width = Number(ext.attributes["cx"] ?? 0);
      result.height = Number(ext.attributes["cy"] ?? 0);
    }

    return result;
  },
};

// ── GroupTransform2D descriptor ──

export const groupTransform2DDesc: CustomDescriptor<GroupTransform2DOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const base = transform2DDesc.stringify(opts, ctx) ?? "<a:xfrm/>";

    // Append chOff and chExt children before closing tag
    const chOff = `<a:chOff x="${opts.childOffsetX ?? 0}" y="${opts.childOffsetY ?? 0}"/>`;
    const chExt = `<a:chExt cx="${opts.childExtentWidth ?? 0}" cy="${opts.childExtentHeight ?? 0}"/>`;

    if (base.endsWith("/>")) {
      return `<a:xfrm>${chOff}${chExt}</a:xfrm>`;
    }
    return base.replace(/<\/a:xfrm>$/, `${chOff}${chExt}</a:xfrm>`);
  },
  parse(el, ctx) {
    const result = transform2DDesc.parse(el, ctx) as Partial<GroupTransform2DOptions>;

    const chOff = findChild(el, "a:chOff");
    if (chOff?.attributes) {
      result.childOffsetX = Number(chOff.attributes["x"] ?? 0);
      result.childOffsetY = Number(chOff.attributes["y"] ?? 0);
    }

    const chExt = findChild(el, "a:chExt");
    if (chExt?.attributes) {
      result.childExtentWidth = Number(chExt.attributes["cx"] ?? 0);
      result.childExtentHeight = Number(chExt.attributes["cy"] ?? 0);
    }

    return result;
  },
};
