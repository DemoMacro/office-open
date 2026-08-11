/**
 * Background (p:bg) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor, ReadContext } from "@office-open/core/descriptor";
import { parse as coreParse } from "@office-open/core/descriptor";
import {
  createEffectList,
  effectListDesc,
  fillDesc,
  type EffectListOptions,
} from "@office-open/core/drawingml";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { buildFill } from "@shared/drawingml/fill";
import type { FillOptions } from "@shared/drawingml/fill";

// ── Types ──

export interface BackgroundDescriptorOptions {
  fill?: FillOptions;
  effects?: EffectListOptions;
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

  const fillXml = buildFill(opts.fill ?? { type: "none" });

  let effectsXml = "";
  if (opts.effects) {
    effectsXml = createEffectList(opts.effects);
  }

  return `<p:bg${bgAttrs.join("")}><p:bgPr${bgPrAttrs.join("")}>${fillXml}${effectsXml}</p:bgPr></p:bg>`;
}

// ── Parse ──

function parseBackground(el: XmlElement, ctx: ReadContext): BackgroundDescriptorOptions {
  const result: Partial<BackgroundDescriptorOptions> = {};

  if (el.attributes?.["p:bwMode"]) {
    result.blackWhiteMode = String(
      el.attributes["p:bwMode"],
    ) as BackgroundDescriptorOptions["blackWhiteMode"];
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

    // Effects (a:effectLst)
    const effectLst = findChild(bgPr, "a:effectLst");
    if (effectLst) {
      const effects = coreParse(effectListDesc, effectLst, ctx);
      if (effects && Object.keys(effects).length > 0) result.effects = effects;
    }
  }

  return result as BackgroundDescriptorOptions;
}
