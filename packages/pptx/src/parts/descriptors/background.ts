/**
 * Background (p:bg) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as coreParse } from "@office-open/core/descriptor";
import { createEffectList, effectListDesc, fillDesc } from "@office-open/core/drawing";
import { stringifyColorChoice, parseColorChoiceElement } from "@office-open/core/drawing";
import { attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { buildFill } from "@shared/drawing/fill";

import type { BackgroundOptions, StyleMatrixReferenceOptions } from "../background";

// ── Descriptor ──

export const backgroundDesc: CustomDescriptor<BackgroundOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    return stringifyBackgroundInner(opts, ctx);
  },

  parse(el, ctx) {
    return parseBackground(el, ctx);
  },
};

// ── Stringify ──

function stringifyBackgroundInner(opts: BackgroundOptions, ctx: WriteContext): string {
  // CT_Background declares an unprefixed @bwMode (a:ST_BlackWhiteMode).
  const bgAttrs: string[] = [];
  if (opts.blackWhiteMode) bgAttrs.push(` bwMode="${opts.blackWhiteMode}"`);

  // CT_Background is a choice: p:bgRef (theme style index) or p:bgPr (explicit fill).
  if (opts.reference) {
    const colorXml =
      opts.reference.color !== undefined ? stringifyColorChoice(opts.reference.color, ctx) : "";
    return `<p:bg${bgAttrs.join("")}><p:bgRef idx="${opts.reference.index}">${colorXml}</p:bgRef></p:bg>`;
  }

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

  if (el.attributes?.["bwMode"]) {
    result.blackWhiteMode = String(el.attributes["bwMode"]) as BackgroundOptions["blackWhiteMode"];
  }

  const bgRef = findChild(el, "p:bgRef");
  if (bgRef) {
    const reference: StyleMatrixReferenceOptions = { index: attrNum(bgRef, "idx") ?? 0 };
    for (const child of bgRef.elements ?? []) {
      const color = parseColorChoiceElement(child, ctx);
      if (color) {
        reference.color = color;
        break;
      }
    }
    result.reference = reference;
    return result as BackgroundOptions;
  }

  const bgPr = findChild(el, "p:bgPr");
  if (bgPr) {
    if (parseOnOff(bgPr.attributes?.["shadeToTitle"])) {
      result.shadeToTitle = true;
    }

    // Fill — CT_BackgroundProperties requires an EG_FillProperties child, so
    // parse always yields one ({ type: "none" } for a:noFill).
    result.fill = coreParse(fillDesc, bgPr, ctx);

    // Effects (a:effectLst) — an empty element stays (bare <a:effectLst/>).
    const effectLst = findChild(bgPr, "a:effectLst");
    if (effectLst) {
      result.effects = coreParse(effectListDesc, effectLst, ctx);
    }
  }

  return result as BackgroundOptions;
}
