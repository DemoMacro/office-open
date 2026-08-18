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

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { parse, stringify } from "../../descriptor";
import { emitPercent, parsePercentAttr } from "../../util/converters";
import { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle } from "../../util/mappings";
import { parseOnOff } from "../../util/values";
import { parseColorChoice, stringifyColorChoice } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import { effectListDesc } from "../effects/effect-descriptors";
import type { EffectListOptions } from "../effects/effect-list";
import { fillDesc, findFillChild } from "../fill/fill-descriptors";
import type { FillOptions } from "../fill/fill-options";
import type { OutlineOptions } from "../outline/outline";
import { outlineDesc, stringifyLineProperties } from "../outline/outline-descriptors";
import {
  DEFAULT_OUTLINE_WIDTH,
  DEFAULT_SHADOW_ALPHA,
  DEFAULT_SHADOW_BLUR_RADIUS,
  DEFAULT_SHADOW_DIRECTION,
  DEFAULT_SHADOW_DISTANCE,
} from "./types";
import type { TextHyperlinkOptions, TextCharacterPropertiesOptions, TextFont } from "./types";

/** Strip `readonly` from all properties. */
export type Mutable<T> = { -readonly [K in keyof T]?: T[K] };

// ── Hyperlink ID counter (module-scoped) ──

let nextHyperlinkId = 1;

/**
 * Register a hyperlink target on the write context and return its placeholder
 * key ("{hlink:key}" on the wire). Action-only hyperlinks (no url/slide) carry
 * no relationship — returns undefined. Shared by text runs and shape cNvPr
 * hyperlinks (both serialize CT_Hyperlink elements).
 */
export function registerHyperlink(hl: TextHyperlinkOptions, ctx: WriteContext): string | undefined {
  if (hl.url === undefined && hl.slide === undefined) return undefined;
  const key = hl.referenceId ?? `hlink_${nextHyperlinkId++}`;
  ctx.addHyperlink(key, { url: hl.url, slide: hl.slide, tooltip: hl.tooltip });
  return key;
}

export function buildHyperlinkElement(
  tag: string,
  hl: TextHyperlinkOptions,
  key: string | undefined,
): string {
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

export function readHyperlink(el: XmlElement, ctx: ReadContext): TextHyperlinkOptions {
  const hl: Mutable<TextHyperlinkOptions> = {};
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
  return hl as TextHyperlinkOptions;
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

/**
 * Serialize CT_TextCharacterProperties under a caller-chosen root tag
 * (a:rPr / a:defRPr / a:endParaRPr — same content model).
 */
export function stringifyRunProperties(
  tag: string,
  opts: TextCharacterPropertiesOptions,
  ctx: WriteContext,
): string {
  // Side-effect: register hyperlinks (click + hover). Only relational targets
  // (url/slide) need a relationship; action-only hyperlinks carry no r:id.
  let hyperlinkKey: string | undefined;
  if (opts.hyperlink) {
    hyperlinkKey = registerHyperlink(opts.hyperlink, ctx);
  }
  let mouseoverKey: string | undefined;
  if (opts.mouseoverHyperlink) {
    mouseoverKey = registerHyperlink(opts.mouseoverHyperlink, ctx);
  }

  const attrParts: string[] = [];
  if (opts.size) attrParts.push(`sz="${Math.round(opts.size * 100)}"`);
  if (opts.bold !== undefined) attrParts.push(`b="${opts.bold ? 1 : 0}"`);
  if (opts.italic !== undefined) attrParts.push(`i="${opts.italic ? 1 : 0}"`);
  if (opts.underline) attrParts.push(`u="${xsdUnderlineStyle.to(opts.underline)}"`);
  if (opts.lang) attrParts.push(`lang="${opts.lang}"`);
  if (opts.strike) attrParts.push(`strike="${xsdStrikeStyle.to(opts.strike)}"`);
  if (opts.baseline !== undefined) attrParts.push(`baseline="${emitPercent(opts.baseline)}"`);
  if (opts.capitalization) attrParts.push(`cap="${xsdTextCaps.to(opts.capitalization)}"`);
  if (opts.spacing !== undefined) attrParts.push(`spc="${Math.round(opts.spacing * 100)}"`);
  if (opts.kern !== undefined) attrParts.push(`kern="${Math.round(opts.kern * 100)}"`);
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

  // EG_TextUnderlineLine / EG_TextUnderlineFill — after highlight, before fonts.
  if (opts.underlineLine !== undefined) {
    if (opts.underlineLine === true) {
      parts.push("<a:uLnTx/>");
    } else {
      // a:uLn is itself a CT_LineProperties — same content, different root tag.
      const uLnXml = stringifyLineProperties("a:uLn", opts.underlineLine, ctx);
      if (uLnXml) parts.push(uLnXml);
    }
  }
  if (opts.underlineFill !== undefined) {
    if (opts.underlineFill === true) {
      parts.push("<a:uFillTx/>");
    } else {
      const uFillXml = stringify(fillDesc, opts.underlineFill, ctx);
      if (uFillXml) parts.push(`<a:uFill>${uFillXml}</a:uFill>`);
    }
  }

  if (opts.font !== undefined) {
    const scripts: {
      latin?: TextFont;
      eastAsia?: TextFont;
      complexScript?: TextFont;
      symbol?: TextFont;
    } = typeof opts.font === "string" ? { latin: opts.font, eastAsia: opts.font } : opts.font;
    if (scripts.latin) parts.push(stringifyFontElement("a:latin", scripts.latin));
    if (scripts.eastAsia) parts.push(stringifyFontElement("a:ea", scripts.eastAsia));
    if (scripts.complexScript) parts.push(stringifyFontElement("a:cs", scripts.complexScript));
    if (scripts.symbol) parts.push(stringifyFontElement("a:sym", scripts.symbol));
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

  if (parts.length === 0) return `<${tag}${attrStr}/>`;
  return `<${tag}${attrStr}>${parts.join("")}</${tag}>`;
}

export const runPropertiesDesc: CustomDescriptor<TextCharacterPropertiesOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    return stringifyRunProperties("a:rPr", opts, ctx);
  },

  parse(el, _ctx) {
    const result: Mutable<TextCharacterPropertiesOptions> = {};

    // nativeTypeAttributes (opc parser) coerces boolean attribute values between
    // string/number/boolean forms; parseOnOff accepts all of them.
    const isOn = (raw: string | number | boolean | undefined): boolean => parseOnOff(raw) ?? false;

    if (el.attributes) {
      if (el.attributes["sz"] !== undefined) result.size = Number(el.attributes["sz"]) / 100;
      if (el.attributes["b"] !== undefined) result.bold = isOn(el.attributes["b"]);
      if (el.attributes["i"] !== undefined) result.italic = isOn(el.attributes["i"]);
      if (el.attributes["u"] !== undefined)
        result.underline = xsdUnderlineStyle.from(
          String(el.attributes["u"]),
        ) as TextCharacterPropertiesOptions["underline"];
      if (el.attributes["lang"] !== undefined) result.lang = String(el.attributes["lang"]);
      if (el.attributes["strike"] !== undefined)
        result.strike = xsdStrikeStyle.from(
          String(el.attributes["strike"]),
        ) as TextCharacterPropertiesOptions["strike"];
      if (el.attributes["baseline"] !== undefined) {
        result.baseline = parsePercentAttr(el.attributes["baseline"])!;
      }
      if (el.attributes["cap"] !== undefined)
        result.capitalization = xsdTextCaps.from(
          String(el.attributes["cap"]),
        ) as TextCharacterPropertiesOptions["capitalization"];
      if (el.attributes["spc"] !== undefined) result.spacing = Number(el.attributes["spc"]) / 100;
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
      if (el.attributes["kern"] !== undefined) result.kern = Number(el.attributes["kern"]) / 100;
      if (el.attributes["err"] !== undefined) result.err = isOn(el.attributes["err"]);
      if (el.attributes["smtClean"] !== undefined)
        result.smtClean = isOn(el.attributes["smtClean"]);
    }

    // Outline — full CT_LineProperties round-trip
    const ln = findChild(el, "a:ln");
    if (ln) result.outline = parse(outlineDesc, ln, _ctx) as OutlineOptions;

    // Fill — CT_TextCharacterProperties references the full EG_FillProperties
    // choice, so accept all six fill kinds (grpFill/blipFill included).
    const fillChild = findFillChild(el);
    if (fillChild) {
      result.fill = parse(fillDesc, fillChild, _ctx) as FillOptions;
    }

    // Shadow — full EG_EffectProperties round-trip
    const effectLst = findChild(el, "a:effectLst");
    if (effectLst) result.shadow = parse(effectListDesc, effectLst, _ctx) as EffectListOptions;

    // Highlight (CT_Color)
    const highlight = findChild(el, "a:highlight");
    if (highlight) result.highlight = parseColorChoice(highlight, _ctx) as SolidFillOptions;

    // Underline line/fill (EG_TextUnderlineLine / EG_TextUnderlineFill)
    const uLnTx = findChild(el, "a:uLnTx");
    if (uLnTx) result.underlineLine = true;
    else {
      const uLn = findChild(el, "a:uLn");
      if (uLn) result.underlineLine = parse(outlineDesc, uLn, _ctx) as OutlineOptions;
    }
    const uFillTx = findChild(el, "a:uFillTx");
    if (uFillTx) result.underlineFill = true;
    else {
      const uFill = findChild(el, "a:uFill");
      if (uFill) result.underlineFill = parse(fillDesc, uFill, _ctx) as FillOptions;
    }

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
      const font: {
        latin?: TextFont;
        eastAsia?: TextFont;
        complexScript?: TextFont;
        symbol?: TextFont;
      } = {};
      if (latinFont) font.latin = latinFont;
      if (eaFont) font.eastAsia = eaFont;
      if (csFont) font.complexScript = csFont;
      if (symFont) font.symbol = symFont;
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

    return result as TextCharacterPropertiesOptions;
  },
};
