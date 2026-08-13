/**
 * RunProperties descriptor (a:rPr / CT_TextCharacterProperties).
 *
 * Promoted from PPTX so all formats share one run-properties model. The
 * hyperlink side-effect uses the core WriteContext.addHyperlink hook — no
 * format-specific cast.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext } from "../../descriptor";
import { parse, stringify } from "../../descriptor";
import { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle } from "../../util/mappings";
import { parseColorChoice, stringifyColorChoice } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import { effectListDesc } from "../effects/effect-descriptors";
import type { EffectListOptions } from "../effects/effect-list";
import { fillDesc } from "../fill/fill-descriptors";
import type { FillOptions } from "../fill/fill-options";
import type { OutlineOptions } from "../outline/outline";
import { outlineDesc } from "../outline/outline-descriptors";
import {
  DEFAULT_OUTLINE_WIDTH,
  DEFAULT_SHADOW_ALPHA,
  DEFAULT_SHADOW_BLUR_RADIUS,
  DEFAULT_SHADOW_DIRECTION,
  DEFAULT_SHADOW_DISTANCE,
} from "./types";
import type { HyperlinkOptions, RunPropertiesOptions, TextFont } from "./types";

/** Strip `readonly` from all properties. */
export type Mutable<T> = { -readonly [K in keyof T]?: T[K] };

// ── Hyperlink ID counter (module-scoped) ──

let nextHyperlinkId = 1;

function buildHyperlinkElement(tag: string, hl: HyperlinkOptions, key: string | undefined): string {
  const attrs: string[] = [];
  // CT_Hyperlink r:id is optional — emit it only for relational targets
  // (external url or internal slide). Action-only tokens (nextslide/endshow/
  // macro/program/...) carry no r:id.
  if ((hl.url !== undefined || hl.slide !== undefined) && key !== undefined) {
    attrs.push(`r:id="{hlink:${key}}"`);
  }
  // Internal slide jump takes precedence over an explicit action token.
  if (hl.slide !== undefined) {
    attrs.push('action="ppaction://hlinksldjump"');
  } else if (hl.action) {
    attrs.push(`action="${escapeXml(hl.action)}"`);
  }
  if (hl.tooltip) attrs.push(`tooltip="${escapeXml(hl.tooltip)}"`);
  if (hl.highlightClick) attrs.push('highlightClick="1"');
  if (hl.endSound) attrs.push('endSnd="1"');
  if (hl.invalidUrl) attrs.push('invalidUrl="1"');
  return attrs.length ? `<${tag} ${attrs.join(" ")}/>` : `<${tag}/>`;
}

function readHyperlink(el: XmlElement, ctx: ReadContext): HyperlinkOptions {
  const hl: Mutable<HyperlinkOptions> = {};
  const rId = el.attributes?.["r:id"];
  const action =
    el.attributes?.["action"] !== undefined ? String(el.attributes["action"]) : undefined;
  if (rId) {
    const ridStr = String(rId);
    const target = ctx.resolveRelationship(ridStr);
    if (target) {
      // Internal slide jump: r:id resolves to slides/slideN.xml.
      const slideMatch = target.match(/slide(\d+)\.xml$/);
      if (slideMatch && action === "ppaction://hlinksldjump") {
        hl.slide = Number(slideMatch[1]);
      } else {
        hl.url = target;
      }
    }
    const m = ridStr.match(/^\{hlink:(.+)\}$/);
    if (m) hl.referenceId = m[1];
  }
  if (el.attributes?.["tooltip"]) hl.tooltip = String(el.attributes["tooltip"]);
  // Preserve explicit action only when it isn't the synthesized slide-jump token
  // (slide already captures that intent).
  if (action && action !== "ppaction://hlinksldjump") hl.action = action;
  if (el.attributes?.["highlightClick"]) hl.highlightClick = true;
  if (el.attributes?.["endSnd"]) hl.endSound = true;
  if (el.attributes?.["invalidUrl"]) hl.invalidUrl = true;
  return hl as HyperlinkOptions;
}

// ── CT_TextFont helpers (a:latin / a:ea / a:cs / a:sym) ──

function stringifyFontElement(tag: string, font: TextFont): string {
  if (typeof font === "string") return `<${tag} typeface="${escapeXml(font)}"/>`;
  const attrs = [`typeface="${escapeXml(font.typeface)}"`];
  if (font.panose) attrs.push(`panose="${font.panose}"`);
  if (font.pitchFamily !== undefined) attrs.push(`pitchFamily="${font.pitchFamily}"`);
  if (font.charset !== undefined) attrs.push(`charset="${font.charset}"`);
  return `<${tag} ${attrs.join(" ")}/>`;
}

function parseFontElement(el: XmlElement | undefined): TextFont | undefined {
  if (!el) return undefined;
  const tf = el.attributes?.["typeface"];
  if (tf === undefined) return undefined;
  const typeface = String(tf);
  const panose = el.attributes?.["panose"];
  const pitchFamily = el.attributes?.["pitchFamily"];
  const charset = el.attributes?.["charset"];
  if (panose === undefined && pitchFamily === undefined && charset === undefined) {
    return typeface;
  }
  const font: { typeface: string; panose?: string; pitchFamily?: number; charset?: number } = {
    typeface,
  };
  if (panose !== undefined) font.panose = String(panose);
  if (pitchFamily !== undefined) font.pitchFamily = Number(pitchFamily);
  if (charset !== undefined) font.charset = Number(charset);
  return font;
}

export const runPropertiesDesc: CustomDescriptor<RunPropertiesOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Side-effect: register hyperlinks (click + hover). Only relational targets
    // (url/slide) need a relationship; action-only hyperlinks carry no r:id.
    let hyperlinkKey: string | undefined;
    if (opts.hyperlink) {
      const hl = opts.hyperlink;
      if (hl.url !== undefined || hl.slide !== undefined) {
        hyperlinkKey = hl.referenceId ?? `hlink_${nextHyperlinkId++}`;
        ctx.addHyperlink(hyperlinkKey, { url: hl.url, slide: hl.slide, tooltip: hl.tooltip });
      }
    }
    let mouseoverKey: string | undefined;
    if (opts.mouseoverHyperlink) {
      const mhl = opts.mouseoverHyperlink;
      if (mhl.url !== undefined || mhl.slide !== undefined) {
        mouseoverKey = mhl.referenceId ?? `hlink_${nextHyperlinkId++}`;
        ctx.addHyperlink(mouseoverKey, { url: mhl.url, slide: mhl.slide, tooltip: mhl.tooltip });
      }
    }

    const attrParts: string[] = [];
    if (opts.size) attrParts.push(`sz="${opts.size * 100}"`);
    if (opts.bold !== undefined) attrParts.push(`b="${opts.bold ? 1 : 0}"`);
    if (opts.italic !== undefined) attrParts.push(`i="${opts.italic ? 1 : 0}"`);
    if (opts.underline) attrParts.push(`u="${xsdUnderlineStyle.to(opts.underline)}"`);
    if (opts.lang) attrParts.push(`lang="${opts.lang}"`);
    if (opts.strike) attrParts.push(`strike="${xsdStrikeStyle.to(opts.strike)}"`);
    if (opts.baseline !== undefined) attrParts.push(`baseline="${opts.baseline * 1000}"`);
    if (opts.capitalization) attrParts.push(`cap="${xsdTextCaps.to(opts.capitalization)}"`);
    if (opts.spacing !== undefined) attrParts.push(`spc="${opts.spacing}"`);
    if (opts.kern !== undefined) attrParts.push(`kern="${opts.kern}"`);
    if (opts.noProof !== undefined) attrParts.push(`noProof="${opts.noProof ? 1 : 0}"`);
    if (opts.dirty !== undefined) attrParts.push(`dirty="${opts.dirty ? 1 : 0}"`);
    if (opts.kumimoji !== undefined) attrParts.push(`kumimoji="${opts.kumimoji ? 1 : 0}"`);
    if (opts.alternateLanguage) attrParts.push(`altLang="${opts.alternateLanguage}"`);
    if (opts.normalizeHeight !== undefined)
      attrParts.push(`normalizeH="${opts.normalizeHeight ? 1 : 0}"`);
    if (opts.bookmarkMark) attrParts.push(`bmk="${opts.bookmarkMark}"`);
    if (opts.smartTagId) attrParts.push(`smtId="${opts.smartTagId}"`);
    if (opts.err !== undefined) attrParts.push(`err="${opts.err ? 1 : 0}"`);
    if (opts.smtClean !== undefined) attrParts.push(`smtClean="${opts.smtClean ? 1 : 0}"`);

    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    const parts: string[] = [];

    // XSD order: ln -> fill -> effect -> latin/ea/cs/sym -> hlinkClick -> rtl
    if (opts.outline !== undefined) {
      const outlineOpts: OutlineOptions =
        opts.outline === true
          ? { width: DEFAULT_OUTLINE_WIDTH, type: "solidFill", color: { value: "000000" } }
          : opts.outline;
      const outlineXml = stringify(outlineDesc, outlineOpts, ctx);
      if (outlineXml) parts.push(outlineXml);
    }

    if (opts.fill !== undefined) {
      const fillXml = stringify(fillDesc, opts.fill, ctx);
      if (fillXml) parts.push(fillXml);
    }

    if (opts.shadow !== undefined) {
      const effectOpts: EffectListOptions =
        opts.shadow === true
          ? {
              outerShadow: {
                blurRadius: DEFAULT_SHADOW_BLUR_RADIUS,
                distance: DEFAULT_SHADOW_DISTANCE,
                direction: DEFAULT_SHADOW_DIRECTION,
                color: { value: "000000", transforms: { alpha: DEFAULT_SHADOW_ALPHA } },
              },
            }
          : opts.shadow;
      const effectXml = stringify(effectListDesc, effectOpts, ctx);
      if (effectXml) parts.push(effectXml);
    }

    // XSD order: highlight (CT_Color) comes after effects, before fonts.
    if (opts.highlight !== undefined) {
      const hlXml = stringifyColorChoice(opts.highlight, ctx);
      if (hlXml) parts.push(`<a:highlight>${hlXml}</a:highlight>`);
    }

    if (opts.font !== undefined) {
      const scripts: { latin?: TextFont; ea?: TextFont; cs?: TextFont; sym?: TextFont } =
        typeof opts.font === "string" ? { latin: opts.font, ea: opts.font } : opts.font;
      if (scripts.latin) parts.push(stringifyFontElement("a:latin", scripts.latin));
      if (scripts.ea) parts.push(stringifyFontElement("a:ea", scripts.ea));
      if (scripts.cs) parts.push(stringifyFontElement("a:cs", scripts.cs));
      if (scripts.sym) parts.push(stringifyFontElement("a:sym", scripts.sym));
    }

    if (opts.hyperlink) {
      parts.push(buildHyperlinkElement("a:hlinkClick", opts.hyperlink, hyperlinkKey));
    }
    if (opts.mouseoverHyperlink) {
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
      if (el.attributes["baseline"] !== undefined) {
        const raw = el.attributes["baseline"];
        const s = typeof raw === "number" ? String(raw) : raw;
        result.baseline = s.endsWith("%") ? Number(s.slice(0, -1)) : Number(s) / 1000;
      }
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
      if (el.attributes["kern"] !== undefined) result.kern = Number(el.attributes["kern"]);
      if (el.attributes["err"] !== undefined) result.err = isOn(el.attributes["err"]);
      if (el.attributes["smtClean"] !== undefined)
        result.smtClean = isOn(el.attributes["smtClean"]);
    }

    // Outline — full CT_LineProperties round-trip
    const ln = findChild(el, "a:ln");
    if (ln) result.outline = parse(outlineDesc, ln, _ctx) as OutlineOptions;

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

    // Shadow — full EG_EffectProperties round-trip
    const effectLst = findChild(el, "a:effectLst");
    if (effectLst) result.shadow = parse(effectListDesc, effectLst, _ctx) as EffectListOptions;

    // Highlight (CT_Color)
    const highlight = findChild(el, "a:highlight");
    if (highlight) result.highlight = parseColorChoice(highlight, _ctx) as SolidFillOptions;

    // Font — latin/ea/cs/sym (CT_TextFont). Collapse to a string when only latin
    // and ea share the same typeface (the common case); otherwise the full object.
    const latinFont = parseFontElement(findChild(el, "a:latin"));
    const eaFont = parseFontElement(findChild(el, "a:ea"));
    const csFont = parseFontElement(findChild(el, "a:cs"));
    const symFont = parseFontElement(findChild(el, "a:sym"));
    if (
      latinFont &&
      eaFont &&
      !csFont &&
      !symFont &&
      typeof latinFont === "string" &&
      typeof eaFont === "string" &&
      latinFont === eaFont
    ) {
      result.font = latinFont;
    } else if (latinFont || eaFont || csFont || symFont) {
      const font: { latin?: TextFont; ea?: TextFont; cs?: TextFont; sym?: TextFont } = {};
      if (latinFont) font.latin = latinFont;
      if (eaFont) font.ea = eaFont;
      if (csFont) font.cs = csFont;
      if (symFont) font.sym = symFont;
      result.font = font;
    }

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
