/**
 * Text element descriptors for PPTX — TextRun (a:r), RunProperties (a:rPr), Paragraph (a:p).
 *
 * @module
 */

import {
  DEFAULT_OUTLINE_WIDTH,
  DEFAULT_SHADOW_ALPHA,
  DEFAULT_SHADOW_BLUR_RADIUS,
  DEFAULT_SHADOW_DIRECTION,
  DEFAULT_SHADOW_DISTANCE,
} from "@file/constants";
import type { FillOptions } from "@file/drawingml/fill";
import type {
  BulletAutoNumOptions,
  BulletCharOptions,
  BulletOptions,
  ParagraphPropertiesOptions,
  TextAlignment,
} from "@file/shape/paragraph/paragraph-properties";
import type { RunOptions } from "@file/shape/paragraph/run";
import type { HyperlinkOptions, RunPropertiesOptions } from "@file/shape/paragraph/run-properties";
import { xsdStrikeStyle, xsdTextAlign, xsdTextCaps, xsdUnderlineStyle } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import { effectListDesc, fillDesc, outlineDesc } from "@office-open/core/drawingml";
import { escapeXml, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { PptxWriteContext } from "../context";

// ── Re-export option types ──

export type { BulletAutoNumOptions, BulletCharOptions, BulletOptions, ParagraphPropertiesOptions };
export type { RunOptions };
export type { HyperlinkOptions, RunPropertiesOptions };

/** Strip `readonly` from all properties. */
type Mutable<T> = { -readonly [K in keyof T]?: T[K] };

export interface ParagraphDescriptorOptions {
  text?: string;
  properties?: ParagraphPropertiesOptions;
  children?: RunOptions[];
}

// ── Hyperlink ID counter (module-scoped) ──

let nextHyperlinkId = 1;

// ── RunProperties (a:rPr) ──

export const runPropertiesDesc: CustomDescriptor<RunPropertiesOptions> = {
  kind: "custom",

  stringify(opts, wctx) {
    const ctx = wctx as PptxWriteContext;

    // Side-effect: register hyperlink
    let hyperlinkKey: string | undefined;
    if (opts.hyperlink) {
      hyperlinkKey = `hlink_${nextHyperlinkId++}`;
      ctx.addHyperlink(hyperlinkKey, opts.hyperlink.url, opts.hyperlink.tooltip);
    }

    const attrParts: string[] = [];
    if (opts.fontSize) attrParts.push(`sz="${opts.fontSize * 100}"`);
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
      const hl = opts.hyperlink;
      const hlAttrs: string[] = [`r:id="{hlink:${hyperlinkKey}}"`];
      if (hl.tooltip) hlAttrs.push(`tooltip="${hl.tooltip}"`);
      if (hl.action) hlAttrs.push(`action="${hl.action}"`);
      if (hl.highlightClick) hlAttrs.push('highlightClick="1"');
      if (hl.endSound) hlAttrs.push('endSnd="1"');
      if (hl.invalidUrl) hlAttrs.push('invalidUrl="1"');
      parts.push(`<a:hlinkClick ${hlAttrs.join(" ")}/>`);
    }

    if (opts.rightToLeft !== undefined) {
      parts.push(`<a:rtl val="${opts.rightToLeft ? 1 : 0}"/>`);
    }

    if (attrParts.length === 0 && parts.length === 0) return "";

    if (parts.length === 0) return `<a:rPr ${attrStr}/>`;
    return `<a:rPr${attrStr}>${parts.join("")}</a:rPr>`;
  },

  parse(el, _ctx) {
    const result: Mutable<RunPropertiesOptions> = {};

    if (el.attributes) {
      if (el.attributes["sz"] !== undefined) result.fontSize = Number(el.attributes["sz"]) / 100;
      if (el.attributes["b"] !== undefined) result.bold = el.attributes["b"] === "1";
      if (el.attributes["i"] !== undefined) result.italic = el.attributes["i"] === "1";
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
      if (el.attributes["noProof"] !== undefined) result.noProof = el.attributes["noProof"] === "1";
      if (el.attributes["dirty"] !== undefined) result.dirty = el.attributes["dirty"] === "1";
      if (el.attributes["kumimoji"] !== undefined)
        result.kumimoji = el.attributes["kumimoji"] === "1";
      if (el.attributes["altLang"] !== undefined)
        result.alternateLanguage = String(el.attributes["altLang"]);
      if (el.attributes["normalizeH"] !== undefined)
        result.normalizeHeight = el.attributes["normalizeH"] === "1";
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

    // Hyperlink
    const hlinkClick = findChild(el, "a:hlinkClick");
    if (hlinkClick) {
      const hl: Mutable<HyperlinkOptions> = {};
      const rId = hlinkClick.attributes?.["r:id"];
      if (rId) {
        const url = _ctx.resolveRelationship(String(rId));
        if (url) hl.url = url;
      }
      if (hlinkClick.attributes?.["tooltip"]) hl.tooltip = String(hlinkClick.attributes["tooltip"]);
      if (hlinkClick.attributes?.["action"]) hl.action = String(hlinkClick.attributes["action"]);
      if (hlinkClick.attributes?.["highlightClick"]) hl.highlightClick = true;
      if (hlinkClick.attributes?.["endSnd"]) hl.endSound = true;
      if (hlinkClick.attributes?.["invalidUrl"]) hl.invalidUrl = true;
      result.hyperlink = hl as HyperlinkOptions;
    }

    // RTL
    const rtl = findChild(el, "a:rtl");
    if (rtl) result.rightToLeft = rtl.attributes?.["val"] === "1";

    return result;
  },
};

// ── TextRun (a:r) ──

export const textRunDesc: CustomDescriptor<RunOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const body = runPropertiesDesc.stringify(opts, ctx) ?? "";

    if (opts.text) {
      return `<a:r>${body}<a:t>${escapeXml(opts.text)}</a:t></a:r>`;
    }
    return body ? `<a:r>${body}</a:r>` : "<a:r/>";
  },

  parse(el, _ctx) {
    const result: Mutable<RunOptions> = {};

    const rPr = findChild(el, "a:rPr");
    if (rPr) {
      Object.assign(result, runPropertiesDesc.parse(rPr, _ctx));
    }

    const t = findChild(el, "a:t");
    if (t) {
      result.text = (t.elements ?? [])
        .filter((e) => e.type === "text")
        .map((e) => e.text ?? "")
        .join("");
    }

    return result;
  },
};

// ── Paragraph properties helper (a:pPr) ──

function stringifyParagraphProperties(options: ParagraphPropertiesOptions): string {
  const children: string[] = [];

  const attrs: string[] = [];
  if (options.alignment) attrs.push(`algn="${xsdTextAlign.to(options.alignment)}"`);
  if (options.indentLevel !== undefined) attrs.push(`lvl="${options.indentLevel}"`);
  if (options.marginIndent !== undefined) attrs.push(`marL="${options.marginIndent}"`);
  if (options.marginRight !== undefined) attrs.push(`marR="${options.marginRight}"`);
  if (options.defTabSize !== undefined) attrs.push(`defTabSz="${options.defTabSize}"`);
  if (options.fontAlignment) attrs.push(`fontAlgn="${options.fontAlignment}"`);

  // Line spacing
  if (options.lineSpacing !== undefined) {
    children.push(`<a:lnSpc><a:spcPct val="${options.lineSpacing * 1000}"/></a:lnSpc>`);
  }
  if (options.lineSpacingPoints !== undefined) {
    children.push(`<a:lnSpc><a:spcPts val="${options.lineSpacingPoints * 100}"/></a:lnSpc>`);
  }

  // Margins
  if (options.marginBottom !== undefined || options.marginTop !== undefined) {
    children.push(`<a:spcAft><a:spcPts val="${options.marginBottom ?? 0}"/></a:spcAft>`);
  }
  if (options.marginTop !== undefined) {
    children.push(`<a:spcBef><a:spcPts val="${options.marginTop}"/></a:spcBef>`);
  }

  // Bullets
  if (options.bullet) {
    children.push(...stringifyBullet(options.bullet));
  }

  if (attrs.length === 0 && children.length === 0) return "";

  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  if (children.length === 0) return `<a:pPr${attrStr}/>`;
  return `<a:pPr${attrStr}>${children.join("")}</a:pPr>`;
}

function stringifyBullet(options: BulletOptions): string[] {
  const parts: string[] = [];

  if (options.type === "none") {
    parts.push("<a:buNone/>");
    return parts;
  }

  if (options.color) {
    parts.push(`<a:buClr><a:srgbClr val="${options.color.replace("#", "")}"/></a:buClr>`);
  }

  if (options.size !== undefined) {
    parts.push(`<a:buSzPct val="${options.size}%"/>`);
  }

  parts.push(
    '<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>',
  );

  if (options.type === "char") {
    parts.push(`<a:buChar char="${escapeXml(options.char ?? "\u2022")}"/>`);
  } else if (options.type === "autoNum") {
    const buAttrs: string[] = [`type="${options.format ?? "arabicPeriod"}"`];
    if (options.startAt !== undefined) buAttrs.push(`startAt="${options.startAt}"`);
    parts.push(`<a:buAutoNum ${buAttrs.join(" ")}/>`);
  }

  return parts;
}

function readParagraphProperties(el: XmlElement): Mutable<ParagraphPropertiesOptions> {
  const result: Mutable<ParagraphPropertiesOptions> = {};

  if (el.attributes) {
    if (el.attributes["algn"] !== undefined)
      result.alignment = xsdTextAlign.from(String(el.attributes["algn"])) as TextAlignment;
    if (el.attributes["lvl"] !== undefined) result.indentLevel = Number(el.attributes["lvl"]);
    if (el.attributes["marL"] !== undefined) result.marginIndent = Number(el.attributes["marL"]);
    if (el.attributes["marR"] !== undefined) result.marginRight = Number(el.attributes["marR"]);
    if (el.attributes["defTabSz"] !== undefined)
      result.defTabSize = Number(el.attributes["defTabSz"]);
    if (el.attributes["fontAlgn"] !== undefined)
      result.fontAlignment = String(
        el.attributes["fontAlgn"],
      ) as ParagraphPropertiesOptions["fontAlignment"];
  }

  // Line spacing
  const lnSpc = findChild(el, "a:lnSpc");
  if (lnSpc) {
    const spcPct = findChild(lnSpc, "a:spcPct");
    if (spcPct?.attributes?.["val"] !== undefined) {
      result.lineSpacing = Number(spcPct.attributes["val"]) / 1000;
    }
    const spcPts = findChild(lnSpc, "a:spcPts");
    if (spcPts?.attributes?.["val"] !== undefined) {
      result.lineSpacingPoints = Number(spcPts.attributes["val"]) / 100;
    }
  }

  // Spacing after/before
  const spcAft = findChild(el, "a:spcAft");
  if (spcAft) {
    const pts = findChild(spcAft, "a:spcPts");
    if (pts?.attributes?.["val"] !== undefined) result.marginBottom = Number(pts.attributes["val"]);
  }

  const spcBef = findChild(el, "a:spcBef");
  if (spcBef) {
    const pts = findChild(spcBef, "a:spcPts");
    if (pts?.attributes?.["val"] !== undefined) result.marginTop = Number(pts.attributes["val"]);
  }

  // Bullets
  if (findChild(el, "a:buNone")) {
    result.bullet = { type: "none" };
  } else if (findChild(el, "a:buChar") || findChild(el, "a:buAutoNum")) {
    const buChar = findChild(el, "a:buChar");
    if (buChar) {
      const bullet: Mutable<BulletCharOptions> = { type: "char" };
      if (buChar.attributes?.["char"]) bullet.char = String(buChar.attributes["char"]);
      const buClr = findChild(el, "a:buClr");
      if (buClr) {
        const srgb = findChild(buClr, "a:srgbClr");
        if (srgb?.attributes?.["val"]) bullet.color = String(srgb.attributes["val"]);
      }
      const buSz = findChild(el, "a:buSzPct");
      if (buSz?.attributes?.["val"])
        bullet.size = Number(String(buSz.attributes["val"]).replace("%", ""));
      result.bullet = bullet as BulletCharOptions;
    } else {
      const buAutoNum = findChild(el, "a:buAutoNum")!;
      const bullet: Mutable<BulletAutoNumOptions> = { type: "autoNum" };
      if (buAutoNum.attributes?.["type"]) bullet.format = String(buAutoNum.attributes["type"]);
      if (buAutoNum.attributes?.["startAt"] !== undefined)
        bullet.startAt = Number(buAutoNum.attributes["startAt"]);
      const buClr = findChild(el, "a:buClr");
      if (buClr) {
        const srgb = findChild(buClr, "a:srgbClr");
        if (srgb?.attributes?.["val"]) bullet.color = String(srgb.attributes["val"]);
      }
      const buSz = findChild(el, "a:buSzPct");
      if (buSz?.attributes?.["val"])
        bullet.size = Number(String(buSz.attributes["val"]).replace("%", ""));
      result.bullet = bullet as BulletAutoNumOptions;
    }
  }

  return result;
}

// ── Paragraph (a:p) ──

export const paragraphDesc: CustomDescriptor<ParagraphDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Paragraph properties
    if (opts.properties) {
      const pPrXml = stringifyParagraphProperties(opts.properties);
      if (pPrXml) parts.push(pPrXml);
    }

    // Simple text shorthand
    if (opts.text) {
      parts.push(textRunDesc.stringify({ text: opts.text }, ctx) ?? "");
    }

    // Children (text runs)
    if (opts.children) {
      for (const child of opts.children) {
        parts.push(textRunDesc.stringify(child, ctx) ?? "");
      }
    }

    parts.push('<a:endParaRPr lang="en-US"/>');

    const body = parts.join("");
    return body ? `<a:p>${body}</a:p>` : "<a:p/>";
  },

  parse(el, ctx) {
    const result: Partial<ParagraphDescriptorOptions> = {};

    // Paragraph properties
    const pPr = findChild(el, "a:pPr");
    if (pPr) {
      result.properties = readParagraphProperties(pPr) as ParagraphPropertiesOptions;
    }

    // Text runs
    const runs: RunOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "a:r") {
        const runOpts = textRunDesc.parse(child, ctx) as RunOptions;
        runs.push(runOpts);
      }
    }

    if (runs.length === 1 && !result.properties) {
      // Single run with no paragraph properties -> use text shorthand
      result.text = runs[0].text;
    } else if (runs.length > 0) {
      result.children = runs;
    }

    return result;
  },
};
