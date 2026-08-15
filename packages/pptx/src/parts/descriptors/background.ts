/**
 * Background (p:bg) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, ReadContext } from "@office-open/core/descriptor";
import { parse as coreParse } from "@office-open/core/descriptor";
import { createEffectList, effectListDesc, fillDesc } from "@office-open/core/drawing";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { buildFill } from "@shared/drawing/fill";

import type { BackgroundOptions } from "../background";

// ── Descriptor ──

export const backgroundDesc: CustomDescriptor<BackgroundOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyBackgroundInner(opts);
  },

  parse(el, ctx) {
    return parseBackground(el, ctx);
  },
};

// ── Stringify ──

function stringifyBackgroundInner(opts: BackgroundOptions): string {
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

function parseBackground(el: XmlElement, ctx: ReadContext): BackgroundOptions {
  const result: Partial<BackgroundOptions> = {};

  if (el.attributes?.["p:bwMode"]) {
    result.blackWhiteMode = String(
      el.attributes["p:bwMode"],
    ) as BackgroundOptions["blackWhiteMode"];
  }

  const bgPr = findChild(el, "p:bgPr");
  if (bgPr) {
    if (parseOnOff(bgPr.attributes?.["shadeToTitle"])) {
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

  return result as BackgroundOptions;
}
