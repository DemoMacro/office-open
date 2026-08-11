/**
 * RunProperties descriptor (a:rPr / CT_TextCharacterProperties).
 *
 * Promoted from PPTX so all formats share one run-properties model. The
 * hyperlink side-effect uses the core WriteContext.addHyperlink hook — no
 * format-specific cast.
 *
 * @module
 */

import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext } from "../../descriptor";
import { parse, stringify } from "../../descriptor";
import { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle } from "../../util/mappings";
import { effectListDesc } from "../effects/effect-descriptors";
import { fillDesc } from "../fill/fill-descriptors";
import type { FillOptions } from "../fill/fill-options";
import { outlineDesc } from "../outline/outline-descriptors";
import {
  DEFAULT_OUTLINE_WIDTH,
  DEFAULT_SHADOW_ALPHA,
  DEFAULT_SHADOW_BLUR_RADIUS,
  DEFAULT_SHADOW_DIRECTION,
  DEFAULT_SHADOW_DISTANCE,
} from "./types";
import type { HyperlinkOptions, RunPropertiesOptions } from "./types";

/** Strip `readonly` from all properties. */
export type Mutable<T> = { -readonly [K in keyof T]?: T[K] };

// ── Hyperlink ID counter (module-scoped) ──

let nextHyperlinkId = 1;

function buildHyperlinkElement(tag: string, hl: HyperlinkOptions, key: string): string {
  const attrs: string[] = [`r:id="{hlink:${key}}"`];
  if (hl.tooltip) attrs.push(`tooltip="${hl.tooltip}"`);
  if (hl.action) attrs.push(`action="${hl.action}"`);
  if (hl.highlightClick) attrs.push('highlightClick="1"');
  if (hl.endSound) attrs.push('endSnd="1"');
  if (hl.invalidUrl) attrs.push('invalidUrl="1"');
  return `<${tag} ${attrs.join(" ")}/>`;
}

function readHyperlink(el: XmlElement, ctx: ReadContext): HyperlinkOptions {
  const hl: Mutable<HyperlinkOptions> = {};
  const rId = el.attributes?.["r:id"];
  if (rId) {
    const ridStr = String(rId);
    const url = ctx.resolveRelationship(ridStr);
    if (url) hl.url = url;
    const m = ridStr.match(/^\{hlink:(.+)\}$/);
    if (m) hl.referenceId = m[1];
  }
  if (el.attributes?.["tooltip"]) hl.tooltip = String(el.attributes["tooltip"]);
  if (el.attributes?.["action"]) hl.action = String(el.attributes["action"]);
  if (el.attributes?.["highlightClick"]) hl.highlightClick = true;
  if (el.attributes?.["endSnd"]) hl.endSound = true;
  if (el.attributes?.["invalidUrl"]) hl.invalidUrl = true;
  return hl as HyperlinkOptions;
}

export const runPropertiesDesc: CustomDescriptor<RunPropertiesOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Side-effect: register hyperlinks (click + hover)
    let hyperlinkKey: string | undefined;
    if (opts.hyperlink) {
      hyperlinkKey = opts.hyperlink.referenceId ?? `hlink_${nextHyperlinkId++}`;
      ctx.addHyperlink(hyperlinkKey, opts.hyperlink.url, opts.hyperlink.tooltip);
    }
    let mouseoverKey: string | undefined;
    if (opts.mouseoverHyperlink) {
      mouseoverKey = opts.mouseoverHyperlink.referenceId ?? `hlink_${nextHyperlinkId++}`;
      ctx.addHyperlink(mouseoverKey, opts.mouseoverHyperlink.url, opts.mouseoverHyperlink.tooltip);
    }

    const attrParts: string[] = [];
    if (opts.size) attrParts.push(`sz="${opts.size * 100}"`);
    if (opts.bold !== undefined) attrParts.push(`b="${opts.bold ? 1 : 0}"`);
    if (opts.italic !== undefined) attrParts.push(`i="${opts.italic ? 1 : 0}"`);
    if (opts.underline) attrParts.push(`u="${xsdUnderlineStyle.to(opts.underline)}"`);
    if (opts.lang) attrParts.push(`lang="${opts.lang}"`);
    if (opts.strike) attrParts.push(`strike="${xsdStrikeStyle.to(opts.strike)}"`);
    if (opts.baseline !== undefined) attrParts.push(`baseline="${opts.baseline}"`);
    if (opts.capitalization) attrParts.push(`cap="${xsdTextCaps.to(opts.capitalization)}"`);
    if (opts.spacing !== undefined) attrParts.push(`spc="${opts.spacing}"`);
    if (opts.noProof !== undefined) attrParts.push(`noProof="${opts.noProof ? 1 : 0}"`);
    if (opts.dirty !== undefined) attrParts.push(`dirty="${opts.dirty ? 1 : 0}"`);
    if (opts.kumimoji !== undefined) attrParts.push(`kumimoji="${opts.kumimoji ? 1 : 0}"`);
    if (opts.alternateLanguage) attrParts.push(`altLang="${opts.alternateLanguage}"`);
    if (opts.normalizeHeight !== undefined)
      attrParts.push(`normalizeH="${opts.normalizeHeight ? 1 : 0}"`);
    if (opts.bookmarkMark) attrParts.push(`bmk="${opts.bookmarkMark}"`);
    if (opts.smartTagId) attrParts.push(`smtId="${opts.smartTagId}"`);

    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    const parts: string[] = [];

    // XSD order: ln -> fill -> effect -> latin/ea -> hlinkClick -> rtl
    if (opts.outline) {
      const outlineXml = stringify(
        outlineDesc,
        { width: DEFAULT_OUTLINE_WIDTH, type: "solidFill", color: { value: "000000" } },
        ctx,
      );
      if (outlineXml) parts.push(outlineXml);
    }

    if (opts.fill !== undefined) {
      const fillXml = stringify(fillDesc, opts.fill, ctx);
      if (fillXml) parts.push(fillXml);
    }

    if (opts.shadow) {
      const effectXml = stringify(
        effectListDesc,
        {
          outerShadow: {
            blurRadius: DEFAULT_SHADOW_BLUR_RADIUS,
            distance: DEFAULT_SHADOW_DISTANCE,
            direction: DEFAULT_SHADOW_DIRECTION,
            color: { value: "000000", transforms: { alpha: DEFAULT_SHADOW_ALPHA } },
          },
        },
        ctx,
      );
      if (effectXml) parts.push(effectXml);
    }

    if (opts.font) {
      parts.push(`<a:latin typeface="${opts.font}"/>`);
      parts.push(`<a:ea typeface="${opts.font}"/>`);
    }

    if (opts.hyperlink && hyperlinkKey) {
      parts.push(buildHyperlinkElement("a:hlinkClick", opts.hyperlink, hyperlinkKey));
    }
    if (opts.mouseoverHyperlink && mouseoverKey) {
      parts.push(buildHyperlinkElement("a:hlinkMouseOver", opts.mouseoverHyperlink, mouseoverKey));
    }

    if (opts.rightToLeft !== undefined) {
      parts.push(`<a:rtl val="${opts.rightToLeft ? 1 : 0}"/>`);
    }

    if (attrParts.length === 0 && parts.length === 0) return "";

    if (parts.length === 0) return `<a:rPr${attrStr}/>`;
    return `<a:rPr${attrStr}>${parts.join("")}</a:rPr>`;
  },

  parse(el, _ctx) {
    const result: Mutable<RunPropertiesOptions> = {};

    // nativeTypeAttributes (opc parser) coerces "1"/"0" to numbers, so a strict
    // `=== "1"` check silently fails; normalize via String() before comparing.
    const isOn = (raw: unknown): boolean => String(raw) === "1";

    if (el.attributes) {
      if (el.attributes["sz"] !== undefined) result.size = Number(el.attributes["sz"]) / 100;
      if (el.attributes["b"] !== undefined) result.bold = isOn(el.attributes["b"]);
      if (el.attributes["i"] !== undefined) result.italic = isOn(el.attributes["i"]);
      if (el.attributes["u"] !== undefined)
        result.underline = xsdUnderlineStyle.from(
          String(el.attributes["u"]),
        ) as RunPropertiesOptions["underline"];
      if (el.attributes["lang"] !== undefined) result.lang = String(el.attributes["lang"]);
      if (el.attributes["strike"] !== undefined)
        result.strike = xsdStrikeStyle.from(
          String(el.attributes["strike"]),
        ) as RunPropertiesOptions["strike"];
      if (el.attributes["baseline"] !== undefined)
        result.baseline = Number(el.attributes["baseline"]);
      if (el.attributes["cap"] !== undefined)
        result.capitalization = xsdTextCaps.from(
          String(el.attributes["cap"]),
        ) as RunPropertiesOptions["capitalization"];
      if (el.attributes["spc"] !== undefined) result.spacing = Number(el.attributes["spc"]);
      if (el.attributes["noProof"] !== undefined) result.noProof = isOn(el.attributes["noProof"]);
      if (el.attributes["dirty"] !== undefined) result.dirty = isOn(el.attributes["dirty"]);
      if (el.attributes["kumimoji"] !== undefined)
        result.kumimoji = isOn(el.attributes["kumimoji"]);
      if (el.attributes["altLang"] !== undefined)
        result.alternateLanguage = String(el.attributes["altLang"]);
      if (el.attributes["normalizeH"] !== undefined)
        result.normalizeHeight = isOn(el.attributes["normalizeH"]);
      if (el.attributes["bmk"] !== undefined) result.bookmarkMark = String(el.attributes["bmk"]);
      if (el.attributes["smtId"] !== undefined) result.smartTagId = String(el.attributes["smtId"]);
    }

    // Outline
    if (findChild(el, "a:ln")) result.outline = true;

    // Fill
    if (
      findChild(el, "a:solidFill") ||
      findChild(el, "a:noFill") ||
      findChild(el, "a:gradFill") ||
      findChild(el, "a:pattFill")
    ) {
      const fillResult = parse(fillDesc, el, _ctx);
      result.fill = fillResult as FillOptions;
    }

    // Shadow
    if (findChild(el, "a:effectLst")) result.shadow = true;

    // Font
    const latin = findChild(el, "a:latin");
    if (latin?.attributes?.["typeface"]) result.font = String(latin.attributes["typeface"]);

    // Hyperlinks (click + hover)
    const hlinkClick = findChild(el, "a:hlinkClick");
    if (hlinkClick) result.hyperlink = readHyperlink(hlinkClick, _ctx);
    const hlinkMouseOver = findChild(el, "a:hlinkMouseOver");
    if (hlinkMouseOver) result.mouseoverHyperlink = readHyperlink(hlinkMouseOver, _ctx);

    // RTL
    const rtl = findChild(el, "a:rtl");
    if (rtl) result.rightToLeft = isOn(rtl.attributes?.["val"]);

    return result as RunPropertiesOptions;
  },
};
