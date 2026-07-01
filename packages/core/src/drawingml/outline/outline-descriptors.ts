/**
 * Outline descriptor for DrawingML shapes.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { convertToEmu } from "../../util/converters";
import { xsdCompoundLine, xsdLineCap, xsdPenAlignment } from "../../util/mappings";
import { solidFillDesc } from "../color/color-descriptors";
import { gradientFillDesc } from "../fill/fill-descriptors";
import type { DashStop } from "./custom-dash";
import type { LineEndOptions } from "./line-end";
import type { OutlineOptions } from "./outline";

// ── LineEnd helpers ──

function stringifyLineEnd(tag: string, opts: LineEndOptions): string {
  const parts: string[] = [];
  if (opts.type) parts.push(`type="${escapeXml(opts.type)}"`);
  if (opts.width) parts.push(`w="${escapeXml(opts.width)}"`);
  if (opts.length) parts.push(`len="${escapeXml(opts.length)}"`);
  const attrStr = parts.length ? " " + parts.join(" ") : "";
  return `<${tag}${attrStr}/>`;
}

function readLineEnd(el: XmlElement): LineEndOptions {
  const result: Partial<LineEndOptions> = {};
  if (el.attributes?.["type"])
    result.type = String(el.attributes["type"]) as LineEndOptions["type"];
  if (el.attributes?.["w"]) result.width = String(el.attributes["w"]) as LineEndOptions["width"];
  if (el.attributes?.["len"])
    result.length = String(el.attributes["len"]) as LineEndOptions["length"];
  return result as LineEndOptions;
}

// ── Custom dash helper ──

function stringifyCustomDash(stops: readonly DashStop[]): string {
  const inner = stops.map((s) => `<a:ds d="${escapeXml(s.d)}" sp="${escapeXml(s.sp)}"/>`).join("");
  return `<a:custDash>${inner}</a:custDash>`;
}

// ── Outline descriptor ──

export const outlineDesc: CustomDescriptor<OutlineOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];

    // Attributes
    const attrParts: string[] = [];
    // ST_LineWidth is EMU (integer) or a universal measure; normalize so the
    // descriptor path matches createOutline (outline.ts) and stays XSD-valid.
    if (opts.width !== undefined) attrParts.push(`w="${convertToEmu(opts.width)}"`);
    // Map full-word enums to XSD tokens (rnd/sng/ctr), matching createOutline.
    if (opts.cap !== undefined) attrParts.push(`cap="${escapeXml(xsdLineCap.to(opts.cap))}"`);
    if (opts.compoundLine !== undefined)
      attrParts.push(`cmpd="${escapeXml(xsdCompoundLine.to(opts.compoundLine))}"`);
    if (opts.align !== undefined)
      attrParts.push(`algn="${escapeXml(xsdPenAlignment.to(opts.align))}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    // Fill
    if (opts.type === "noFill") {
      parts.push("<a:noFill/>");
    } else if (opts.type === "solidFill" && opts.color) {
      const fillXml = stringify(solidFillDesc, opts.color, ctx);
      if (fillXml) parts.push(fillXml);
    } else if (opts.type === "gradFill" && opts.gradientFill) {
      const gradXml = stringify(gradientFillDesc, opts.gradientFill, ctx);
      if (gradXml) parts.push(gradXml);
    }

    // Dash
    if (opts.customDash) {
      parts.push(stringifyCustomDash(opts.customDash));
    } else if (opts.dash) {
      parts.push(`<a:prstDash val="${escapeXml(opts.dash)}"/>`);
    }

    // Join
    if (opts.join) {
      if (opts.join === "miter" && opts.miterLimit !== undefined) {
        parts.push(`<a:miter lim="${opts.miterLimit}"/>`);
      } else {
        parts.push(`<a:${opts.join}/>`);
      }
    }

    // Line ends
    if (opts.headEnd) parts.push(stringifyLineEnd("a:headEnd", opts.headEnd));
    if (opts.tailEnd) parts.push(stringifyLineEnd("a:tailEnd", opts.tailEnd));

    if (parts.length === 0 && !attrStr) return undefined;
    if (parts.length === 0) return `<a:ln${attrStr}/>`;
    return `<a:ln${attrStr}>${parts.join("")}</a:ln>`;
  },
  parse(el, _ctx) {
    const result: OutlineOptions = {};

    // Attributes
    if (el.attributes) {
      if (el.attributes["w"] !== undefined) result.width = Number(el.attributes["w"]);
      // Reverse-map XSD tokens back to full-word enums so the parsed Options
      // match the user-facing type (round/single/center), not raw tokens.
      if (el.attributes["cap"] !== undefined)
        result.cap = xsdLineCap.from(String(el.attributes["cap"])) as OutlineOptions["cap"];
      if (el.attributes["cmpd"] !== undefined)
        result.compoundLine = xsdCompoundLine.from(
          String(el.attributes["cmpd"]),
        ) as OutlineOptions["compoundLine"];
      if (el.attributes["algn"] !== undefined)
        result.align = xsdPenAlignment.from(
          String(el.attributes["algn"]),
        ) as OutlineOptions["align"];
    }

    // Fill
    const solidFill = findChild(el, "a:solidFill");
    if (solidFill) {
      result.type = "solidFill";
      result.color = parse(solidFillDesc, solidFill, _ctx);
    }
    if (findChild(el, "a:noFill")) {
      result.type = "noFill";
    }
    if (findChild(el, "a:gradFill")) {
      result.type = "gradFill";
      const gradFill = findChild(el, "a:gradFill")!;
      result.gradientFill = parse(gradientFillDesc, gradFill, _ctx);
    }

    // Dash
    const prstDash = findChild(el, "a:prstDash");
    if (prstDash?.attributes?.["val"])
      result.dash = String(prstDash.attributes["val"]) as OutlineOptions["dash"];
    const custDash = findChild(el, "a:custDash");
    if (custDash?.elements) {
      result.customDash = custDash.elements
        .filter((c) => c.name === "a:ds")
        .map((c) => ({
          d: String(c.attributes?.["d"] ?? ""),
          sp: String(c.attributes?.["sp"] ?? ""),
        }));
    }

    // Join
    if (findChild(el, "a:round")) result.join = "round";
    else if (findChild(el, "a:bevel")) result.join = "bevel";
    else {
      const miter = findChild(el, "a:miter");
      if (miter) {
        result.join = "miter";
        if (miter.attributes?.["lim"]) result.miterLimit = Number(miter.attributes["lim"]);
      }
    }

    // Line ends
    const headEnd = findChild(el, "a:headEnd");
    if (headEnd) result.headEnd = readLineEnd(headEnd);
    const tailEnd = findChild(el, "a:tailEnd");
    if (tailEnd) result.tailEnd = readLineEnd(tailEnd);

    return result;
  },
};
