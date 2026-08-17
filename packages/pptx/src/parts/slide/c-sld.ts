/**
 * Shared cSld sub-block serializers/parsers (CT_CommonSlideData tail children)
 * used by the slide-family serializers — compiler.ts (slide), slide-master.ts,
 * slide-layout.ts, and the slide descriptor. p:hf is attribute-form per
 * CT_HeaderFooter (sldNum/hdr/ftr/dt boolean attributes, no element children);
 * parse additionally tolerates the legacy element-form emission.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";

import type { ControlOptions } from "./slide";

// ── p:custDataLst ──

/** Serialize cSld customer-data references; empty string when there are none. */
export function stringifyCustDataLst(items: { rId: string }[] | undefined): string {
  if (!items || items.length === 0) return "";
  return `<p:custDataLst>${items.map((d) => `<p:custData r:id="${d.rId}"/>`).join("")}</p:custDataLst>`;
}

/** Parse a p:custDataLst element into customer-data references. */
export function parseCustDataLst(el: Element | undefined): { rId: string }[] | undefined {
  if (!el) return undefined;
  const items: { rId: string }[] = [];
  for (const cd of el.elements ?? []) {
    if (cd.name === "p:custData") {
      const rId = attr(cd, "r:id");
      if (rId) items.push({ rId });
    }
  }
  return items.length > 0 ? items : undefined;
}

// ── p:controls ──

/** Serialize cSld controls; empty string when there are none. */
export function stringifyControls(controls: ControlOptions[] | undefined): string {
  if (!controls || controls.length === 0) return "";
  const items = controls
    .map((c) => {
      const attrs: string[] = [];
      if (c.shapeId !== undefined) attrs.push(`spid="${c.shapeId}"`);
      if (c.name) attrs.push(`name="${c.name}"`);
      if (c.showAsIcon) attrs.push('showAsIcon="1"');
      if (c.rId) attrs.push(`r:id="${c.rId}"`);
      if (c.imageWidth !== undefined) attrs.push(`imgW="${c.imageWidth}"`);
      if (c.imageHeight !== undefined) attrs.push(`imgH="${c.imageHeight}"`);
      return `<p:control ${attrs.join(" ")}/>`;
    })
    .join("");
  return `<p:controls>${items}</p:controls>`;
}

/** Parse a p:controls element into control options. */
export function parseControls(el: Element | undefined): ControlOptions[] | undefined {
  if (!el) return undefined;
  const items: ControlOptions[] = [];
  for (const ctrl of el.elements ?? []) {
    if (ctrl.name !== "p:control") continue;
    const item: ControlOptions = {};
    // @spid is ST_ShapeID (a token like "_x0000_s1026"), never numeric.
    const shapeId = attr(ctrl, "spid");
    if (shapeId) item.shapeId = shapeId;
    const name = attr(ctrl, "name");
    if (name) item.name = name;
    if (parseOnOff(attr(ctrl, "showAsIcon"))) item.showAsIcon = true;
    const rId = attr(ctrl, "r:id");
    if (rId) item.rId = rId;
    const imageWidth = attrNum(ctrl, "imgW");
    if (imageWidth !== undefined) item.imageWidth = imageWidth;
    const imageHeight = attrNum(ctrl, "imgH");
    if (imageHeight !== undefined) item.imageHeight = imageHeight;
    items.push(item);
  }
  return items.length > 0 ? items : undefined;
}

// ── p:hf (CT_HeaderFooter, attribute form) ──

/**
 * Serialize a layout-level p:hf in attribute form per CT_HeaderFooter;
 * undefined/empty options emit nothing (each attribute defaults to true in
 * the XSD, so an all-true header/footer is the absent element).
 */
export function stringifySlideHf(hf: SlideHeaderFooterOptions | undefined): string {
  if (!hf) return "";
  const attrs: string[] = [];
  if (hf.slideNumber !== undefined) attrs.push(`sldNum="${hf.slideNumber ? 1 : 0}"`);
  if (hf.dateTime !== undefined) attrs.push(`dt="${hf.dateTime ? 1 : 0}"`);
  if (hf.footer !== undefined) attrs.push(`ftr="${hf.footer ? 1 : 0}"`);
  if (hf.header !== undefined) attrs.push(`hdr="${hf.header ? 1 : 0}"`);
  if (attrs.length === 0) return "";
  return `<p:hf ${attrs.join(" ")}/>`;
}

/**
 * Parse a layout-level p:hf. Reads CT_HeaderFooter attributes; tolerates the
 * legacy element-form children this library used to emit.
 */
export function parseSlideHf(el: Element | undefined): SlideHeaderFooterOptions | undefined {
  if (!el) return undefined;
  const hf: SlideHeaderFooterOptions = {};
  const sldNum = attr(el, "sldNum");
  if (sldNum !== undefined) hf.slideNumber = parseOnOff(sldNum) ?? false;
  const dt = attr(el, "dt");
  if (dt !== undefined) hf.dateTime = parseOnOff(dt) ?? false;
  const ftr = attr(el, "ftr");
  if (ftr !== undefined) hf.footer = parseOnOff(ftr) ?? false;
  const hdr = attr(el, "hdr");
  if (hdr !== undefined) hf.header = parseOnOff(hdr) ?? false;
  if (Object.keys(hf).length > 0) return hf;

  // Legacy element-form emission (<p:hf><p:sldNum/>…</p:hf>).
  if (findChild(el, "p:sldNum")) hf.slideNumber = true;
  if (findChild(el, "p:dt")) hf.dateTime = true;
  if (findChild(el, "p:ftr")) hf.footer = true;
  if (findChild(el, "p:hdr")) hf.header = true;
  return Object.keys(hf).length > 0 ? hf : undefined;
}
