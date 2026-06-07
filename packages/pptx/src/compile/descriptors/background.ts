/**
 * Background (p:bg) descriptor for PPTX.
 *
 * @module
 */

import { createPptxEffectList } from "@file/drawingml/effects";
import type { EffectsOptions } from "@file/drawingml/effects";
import { buildFill } from "@file/drawingml/fill";
import type { FillOptions } from "@file/drawingml/fill";
import type { CustomDescriptor, ReadContext } from "@office-open/core/descriptor";
import { parse as coreParse } from "@office-open/core/descriptor";
import { fillDesc } from "@office-open/core/drawingml";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

// ── Types ──

export interface BackgroundDescriptorOptions {
  fill?: FillOptions;
  effects?: EffectsOptions;
  shadeToTitle?: boolean;
  blackWhiteMode?:
    | "clr"
    | "gray"
    | "ltGray"
    | "invGray"
    | "gmGray"
    | "bw"
    | "auto"
    | "black"
    | "white";
}

// ── Descriptor ──

export const backgroundDesc: CustomDescriptor<BackgroundDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyBackgroundInner(opts);
  },

  parse(el, ctx) {
    return parseBackground(el, ctx);
  },
};

// ── Stringify ──

function stringifyBackgroundInner(opts: BackgroundDescriptorOptions): string {
  const bgAttrs: string[] = [];
  if (opts.blackWhiteMode) bgAttrs.push(` p:bwMode="${opts.blackWhiteMode}"`);

  const bgPrAttrs: string[] = [];
  if (opts.shadeToTitle) bgPrAttrs.push(' shadeToTitle="1"');

  const fillComp = buildFill(opts.fill ?? { type: "none" });
  const fillXml = fillComp.serialize();

  let effectsXml = "";
  if (opts.effects) {
    const el = createPptxEffectList(opts.effects);
    if (el) effectsXml = el.serialize();
  }

  return `<p:bg${bgAttrs.join("")}><p:bgPr${bgPrAttrs.join("")}>${fillXml}${effectsXml}</p:bgPr></p:bg>`;
}

// ── Parse ──

function parseBackground(el: XmlElement, ctx: ReadContext): Partial<BackgroundDescriptorOptions> {
  const result: Record<string, unknown> = {};

  if (el.attributes?.["p:bwMode"]) {
    result.blackWhiteMode = el.attributes["p:bwMode"];
  }

  const bgPr = findChild(el, "p:bgPr");
  if (bgPr) {
    if (bgPr.attributes?.["shadeToTitle"] === "1") {
      result.shadeToTitle = true;
    }

    // Fill
    const fillResult = coreParse(fillDesc, bgPr, ctx);
    if (fillResult && Object.keys(fillResult).length > 0) {
      result.fill = fillResult;
    }
  }

  return result as Partial<BackgroundDescriptorOptions>;
}
