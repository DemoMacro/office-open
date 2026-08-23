/**
 * Paragraph descriptor (a:p / CT_TextParagraph) plus paragraph-properties,
 * text-field, and line-break helpers.
 *
 * Promoted verbatim from PPTX — the fresh-paragraph default lang marker,
 * a:lstStyle-always-emit, and single-run text shorthand behaviors are
 * load-bearing for round-trip fidelity and must carry over unchanged.
 *
 * @module
 */

import { findChild, escapeXml, stringify as stringifyXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { emitPercent, parsePercent } from "../../util/converters";
import { xsdFontAlignment, xsdTextAlign, xsdTextTabAlignment } from "../../util/mappings";
import { parseOnOff, stripColorHashPrefix } from "../../util/values";
import { emitColorChoice, isPlainRgbColor, parseColorChoice } from "../color/color-descriptors";
import { stringifyTextRun, textRunDesc } from "./run";
import { runPropertiesDesc, stringifyRunProperties } from "./run-properties";
import type { Mutable } from "./run-properties";
import type {
  BulletAutoNumOptions,
  BulletCharOptions,
  BulletOptions,
  BulletPictureOptions,
  BulletStyleOptions,
  TextParagraphPropertiesOptions,
  TextRunOptions,
  TextCharacterPropertiesOptions,
  TabStopOptions,
  TextAlignment,
  TextTabAlignment,
  TextFieldOptions,
  TextBreakOptions,
} from "./types";

/** Descriptor-layer paragraph accumulator (a:p). */
export interface ParagraphDescriptorOptions {
  text?: string;
  properties?: TextParagraphPropertiesOptions;
  children?: (TextRunOptions | TextFieldOptions | TextBreakOptions | string)[];
  /**
   * End-paragraph run properties (a:endParaRPr). Fresh paragraphs emit a
   * default lang marker; a parsed source preserves its value; false omits it.
   */
  endParagraphProperties?: TextCharacterPropertiesOptions | false;
}

// ── Paragraph properties helper (a:pPr / a:lvlNpPr / a:defPPr) ──

/**
 * Serialize CT_TextParagraphProperties under a caller-chosen root tag. The
 * same content model backs a:pPr in paragraphs and a:lvlNpPr/a:defPPr in
 * list styles. Attribute order follows the XSD declaration order so MS
 * Office-authored bytes re-emit unchanged.
 */
export function stringifyParagraphPropertiesElement(
  tag: string,
  options: TextParagraphPropertiesOptions,
  ctx: WriteContext,
): string {
  const children: string[] = [];

  // XSD attribute order: marL, marR, lvl, indent, algn, defTabSz, rtl,
  // eaLnBrk, fontAlgn, latinLnBrk, hangingPunct.
  const attrs: string[] = [];
  if (options.marginIndent !== undefined) attrs.push(`marL="${options.marginIndent}"`);
  if (options.marginRight !== undefined) attrs.push(`marR="${options.marginRight}"`);
  if (options.indentLevel !== undefined) attrs.push(`lvl="${options.indentLevel}"`);
  if (options.indent !== undefined) attrs.push(`indent="${options.indent}"`);
  if (options.alignment) attrs.push(`algn="${xsdTextAlign.to(options.alignment)}"`);
  if (options.defTabSize !== undefined) attrs.push(`defTabSz="${options.defTabSize}"`);
  if (options.rightToLeft !== undefined) attrs.push(`rtl="${options.rightToLeft ? 1 : 0}"`);
  if (options.eastAsianLineBreak !== undefined)
    attrs.push(`eaLnBrk="${options.eastAsianLineBreak ? 1 : 0}"`);
  if (options.fontAlignment) attrs.push(`fontAlgn="${xsdFontAlignment.to(options.fontAlignment)}"`);
  if (options.latinLineBreak !== undefined)
    attrs.push(`latinLnBrk="${options.latinLineBreak ? 1 : 0}"`);
  if (options.hangingPunctuation !== undefined)
    attrs.push(`hangingPunct="${options.hangingPunctuation ? 1 : 0}"`);

  // Line spacing — a percentage and a points value are mutually exclusive
  // (a:lnSpc carries a single spcPct | spcPts child).
  if (options.lineSpacingPercent !== undefined) {
    children.push(
      `<a:lnSpc><a:spcPct val="${emitPercent(options.lineSpacingPercent)}"/></a:lnSpc>`,
    );
  } else if (options.lineSpacingPoints !== undefined) {
    children.push(
      `<a:lnSpc><a:spcPts val="${Math.round(options.lineSpacingPoints * 100)}"/></a:lnSpc>`,
    );
  }

  // Spacing before/after (XSD CT_TextParagraphProperties order: spcBef, spcAft).
  // A percentage and a points value are mutually exclusive per XSD (spcPct | spcPts).
  if (options.spaceBeforePercent !== undefined) {
    children.push(
      `<a:spcBef><a:spcPct val="${emitPercent(options.spaceBeforePercent)}"/></a:spcBef>`,
    );
  } else if (options.spaceBefore !== undefined) {
    children.push(
      `<a:spcBef><a:spcPts val="${Math.round(options.spaceBefore * 100)}"/></a:spcBef>`,
    );
  }
  if (options.spaceAfterPercent !== undefined) {
    children.push(
      `<a:spcAft><a:spcPct val="${emitPercent(options.spaceAfterPercent)}"/></a:spcAft>`,
    );
  } else if (options.spaceAfter !== undefined) {
    children.push(`<a:spcAft><a:spcPts val="${Math.round(options.spaceAfter * 100)}"/></a:spcAft>`);
  }

  // Bullets
  if (options.bullet) {
    children.push(...stringifyBullet(options.bullet));
  }

  // Tab stops (after bullets, before defRPr)
  if (options.tabStops) {
    children.push(stringifyTabStops(options.tabStops));
  }

  // Default run properties — after tabLst (CT_TextParagraphProperties).
  // Presence-based: an explicitly empty object still emits <a:defRPr/> (the
  // otherStyle defPPr default carries one).
  if (options.defaultRunProperties !== undefined) {
    const rPr = stringifyRunProperties("a:defRPr", options.defaultRunProperties, ctx);
    children.push(rPr || "<a:defRPr/>");
  }

  // a:extLst — last child.
  if (options.ext) children.push(`<a:extLst>${options.ext}</a:extLst>`);

  if (attrs.length === 0 && children.length === 0) return "";

  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  if (children.length === 0) return `<${tag}${attrStr}/>`;
  return `<${tag}${attrStr}>${children.join("")}</${tag}>`;
}

function stringifyParagraphProperties(
  options: TextParagraphPropertiesOptions,
  ctx: WriteContext,
): string {
  return stringifyParagraphPropertiesElement("a:pPr", options, ctx);
}

function stringifyBullet(options: BulletOptions): string[] {
  const parts: string[] = [];

  // Color: buClrTx | buClr
  if (options.colorFollowsText) {
    parts.push("<a:buClrTx/>");
  } else if (options.color) {
    const color =
      typeof options.color === "string"
        ? emitColorChoice({ value: stripColorHashPrefix(options.color) })
        : emitColorChoice(options.color);
    parts.push(`<a:buClr>${color}</a:buClr>`);
  }

  // Size: buSzTx | buSzPts | buSzPct
  if (options.sizeFollowsText) {
    parts.push("<a:buSzTx/>");
  } else if (options.sizePoints !== undefined) {
    parts.push(`<a:buSzPts val="${Math.round(options.sizePoints * 100)}"/>`);
  } else if (options.size !== undefined) {
    // ST_TextBulletSizePercent is an "N%" string per the XSD pattern (25-400);
    // the union's per-mille-integer branch is what Office writes but the XSD
    // loader here only enforces the pattern, so emit the string form.
    parts.push(`<a:buSzPct val="${Math.round(options.size)}%"/>`);
  }

  // Font: buFontTx | buFont. No fresh default — Office files omit buFont when
  // the bullet follows the paragraph font, and a synthesized one would not
  // round-trip.
  if (options.fontFollowsText) {
    parts.push("<a:buFontTx/>");
  } else if (options.font !== undefined) {
    parts.push(
      `<a:buFont typeface="${options.font}" panose="020B0604020202020204" pitchFamily="34" charset="0"/>`,
    );
  }

  // Bullet type: buNone | buChar | buAutoNum | buBlip
  if (options.type === "none") {
    parts.push("<a:buNone/>");
  } else if (options.type === "char") {
    parts.push(`<a:buChar char="${escapeXml(options.char ?? "•")}"/>`);
  } else if (options.type === "autoNum") {
    const buAttrs: string[] = [`type="${options.format ?? "arabicPeriod"}"`];
    if (options.startAt !== undefined) buAttrs.push(`startAt="${options.startAt}"`);
    parts.push(`<a:buAutoNum ${buAttrs.join(" ")}/>`);
  } else if (options.type === "picture") {
    // a:blip @r:embed must name a relationship the host part owns — without
    // one there is no image to bullet with, so the element is skipped
    if (options.embed) {
      parts.push(`<a:buBlip><a:blip r:embed="${options.embed}"/></a:buBlip>`);
    }
  }

  return parts;
}

function stringifyTabStops(stops: TabStopOptions[]): string {
  const tabs = stops.map((t) => {
    const attrs: string[] = [];
    if (t.position !== undefined) attrs.push(`pos="${t.position}"`);
    if (t.alignment) attrs.push(`algn="${xsdTextTabAlignment.to(t.alignment)}"`);
    return `<a:tab${attrs.length ? " " + attrs.join(" ") : ""}/>`;
  });
  if (tabs.length === 0) return "<a:tabLst/>";
  return `<a:tabLst>${tabs.join("")}</a:tabLst>`;
}

export function readParagraphProperties(
  el: XmlElement,
  ctx: ReadContext,
): Mutable<TextParagraphPropertiesOptions> {
  const result: Mutable<TextParagraphPropertiesOptions> = {};

  if (el.attributes) {
    if (el.attributes["algn"] !== undefined)
      result.alignment = xsdTextAlign.from(String(el.attributes["algn"])) as TextAlignment;
    if (el.attributes["lvl"] !== undefined) result.indentLevel = Number(el.attributes["lvl"]);
    if (el.attributes["marL"] !== undefined) result.marginIndent = Number(el.attributes["marL"]);
    if (el.attributes["marR"] !== undefined) result.marginRight = Number(el.attributes["marR"]);
    if (el.attributes["indent"] !== undefined) result.indent = Number(el.attributes["indent"]);
    if (el.attributes["defTabSz"] !== undefined)
      result.defTabSize = Number(el.attributes["defTabSz"]);
    if (el.attributes["fontAlgn"] !== undefined)
      result.fontAlignment = xsdFontAlignment.from(
        String(el.attributes["fontAlgn"]),
      ) as TextParagraphPropertiesOptions["fontAlignment"];
    if (el.attributes["rtl"] !== undefined)
      result.rightToLeft = parseOnOff(el.attributes["rtl"]) ?? false;
    if (el.attributes["eaLnBrk"] !== undefined)
      result.eastAsianLineBreak = parseOnOff(el.attributes["eaLnBrk"]) ?? false;
    if (el.attributes["latinLnBrk"] !== undefined)
      result.latinLineBreak = parseOnOff(el.attributes["latinLnBrk"]) ?? false;
    if (el.attributes["hangingPunct"] !== undefined)
      result.hangingPunctuation = parseOnOff(el.attributes["hangingPunct"]) ?? false;
  }

  // Line spacing
  const lnSpc = findChild(el, "a:lnSpc");
  if (lnSpc) {
    const spcPct = findChild(lnSpc, "a:spcPct");
    if (spcPct?.attributes?.["val"] !== undefined) {
      result.lineSpacingPercent = parsePercent(Number(spcPct.attributes["val"]));
    }
    const spcPts = findChild(lnSpc, "a:spcPts");
    if (spcPts?.attributes?.["val"] !== undefined) {
      result.lineSpacingPoints = Number(spcPts.attributes["val"]) / 100;
    }
  }

  // Spacing after/before — each is a spcPct | spcPts choice
  const spcAft = findChild(el, "a:spcAft");
  if (spcAft) {
    const pct = findChild(spcAft, "a:spcPct");
    if (pct?.attributes?.["val"] !== undefined) {
      result.spaceAfterPercent = parsePercent(Number(pct.attributes["val"]));
    } else {
      const pts = findChild(spcAft, "a:spcPts");
      if (pts?.attributes?.["val"] !== undefined)
        result.spaceAfter = Number(pts.attributes["val"]) / 100;
    }
  }

  const spcBef = findChild(el, "a:spcBef");
  if (spcBef) {
    const pct = findChild(spcBef, "a:spcPct");
    if (pct?.attributes?.["val"] !== undefined) {
      result.spaceBeforePercent = parsePercent(Number(pct.attributes["val"]));
    } else {
      const pts = findChild(spcBef, "a:spcPts");
      if (pts?.attributes?.["val"] !== undefined)
        result.spaceBefore = Number(pts.attributes["val"]) / 100;
    }
  }

  // Bullets
  const buNone = findChild(el, "a:buNone");
  {
    const buChar = findChild(el, "a:buChar");
    const buAutoNum = findChild(el, "a:buAutoNum");
    const buBlip = findChild(el, "a:buBlip");
    if (buChar || buAutoNum || buBlip || buNone) {
      // Shared color/size/font style — each dimension is a choice.
      const style: Mutable<BulletStyleOptions> = {};
      if (findChild(el, "a:buClrTx")) {
        style.colorFollowsText = true;
      } else {
        const buClr = findChild(el, "a:buClr");
        if (buClr) {
          const parsed = parseColorChoice(buClr, ctx);
          if (parsed && Object.keys(parsed).length > 0) {
            style.color = isPlainRgbColor(parsed) ? (parsed as { value: string }).value : parsed;
          }
        }
      }
      if (findChild(el, "a:buSzTx")) {
        style.sizeFollowsText = true;
      } else {
        const buSzPts = findChild(el, "a:buSzPts");
        if (buSzPts?.attributes?.["val"]) {
          style.sizePoints = Number(buSzPts.attributes["val"]) / 100;
        } else {
          const buSzPct = findChild(el, "a:buSzPct");
          if (buSzPct?.attributes?.["val"]) {
            // Per-mille integer ("60000" = 60%) is what Office writes; the
            // XSD's "75%"-string form carries the percentage directly.
            const raw = String(buSzPct.attributes["val"]);
            const num = Number(raw.replace("%", ""));
            style.size = raw.includes("%") ? num : num / 1000;
          }
        }
      }
      if (findChild(el, "a:buFontTx")) {
        style.fontFollowsText = true;
      } else {
        const buFont = findChild(el, "a:buFont");
        if (buFont?.attributes?.["typeface"]) style.font = String(buFont.attributes["typeface"]);
      }

      if (buNone) {
        result.bullet = { type: "none", ...style };
      } else if (buBlip) {
        const blip = findChild(buBlip, "a:blip");
        result.bullet = {
          type: "picture",
          embed: String(blip?.attributes?.["r:embed"] ?? ""),
          ...style,
        } as BulletPictureOptions;
      } else if (buChar) {
        const bullet: Mutable<BulletCharOptions> = { type: "char", ...style };
        if (buChar.attributes?.["char"]) bullet.char = String(buChar.attributes["char"]);
        result.bullet = bullet as BulletCharOptions;
      } else {
        const bullet: Mutable<BulletAutoNumOptions> = { type: "autoNum", ...style };
        if (buAutoNum!.attributes?.["type"])
          bullet.format = String(buAutoNum!.attributes["type"]) as BulletAutoNumOptions["format"];
        if (buAutoNum!.attributes?.["startAt"] !== undefined)
          bullet.startAt = Number(buAutoNum!.attributes["startAt"]);
        result.bullet = bullet as BulletAutoNumOptions;
      }
    }
  }

  // Tab stops (after bullets, before defRPr)
  const tabLst = findChild(el, "a:tabLst");
  if (tabLst) {
    const tabs: TabStopOptions[] = [];
    for (const child of tabLst.elements ?? []) {
      if (child.name === "a:tab") {
        const tab: Mutable<TabStopOptions> = {};
        if (child.attributes?.["pos"] !== undefined) tab.position = Number(child.attributes["pos"]);
        if (child.attributes?.["algn"] !== undefined)
          tab.alignment = xsdTextTabAlignment.from(
            String(child.attributes["algn"]),
          ) as TextTabAlignment;
        tabs.push(tab);
      }
    }
    // Keep an explicit empty element — sources often carry a bare <a:tabLst/>.
    result.tabStops = tabs;
  }

  // Default run properties — after tabLst (CT_TextParagraphProperties).
  const defRPr = findChild(el, "a:defRPr");
  if (defRPr) {
    result.defaultRunProperties = runPropertiesDesc.parse(
      defRPr,
      ctx,
    ) as TextCharacterPropertiesOptions;
  }

  // a:extLst — last child; verbatim inner XML for unmodeled extensions.
  const extLst = findChild(el, "a:extLst");
  if (extLst) {
    const inner = stringifyXml(extLst);
    if (inner) result.ext = inner;
  }

  return result as TextParagraphPropertiesOptions;
}

// ── Paragraph (a:p) ──

export const paragraphDesc: CustomDescriptor<ParagraphDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Paragraph properties
    if (opts.properties) {
      const pPrXml = stringifyParagraphProperties(opts.properties, ctx);
      if (pPrXml) parts.push(pPrXml);
    }

    // Simple text shorthand
    if (opts.text) {
      parts.push(stringifyTextRun(opts.text));
    }

    // Children (text runs + fields + line breaks; a bare string is shorthand
    // for a single text-only run).
    if (opts.children) {
      for (const child of opts.children) {
        if (typeof child === "string") {
          parts.push(stringifyTextRun(child));
        } else if (isTextField(child)) {
          parts.push(stringifyTextField(child, ctx));
        } else if (isBreak(child)) {
          parts.push(stringifyBreak(child, ctx));
        } else {
          parts.push(textRunDesc.stringify(child, ctx) ?? "");
        }
      }
    }

    // End-paragraph run properties — fresh emits a default lang marker; a
    // parsed source preserves its value; false omits the element.
    if (opts.endParagraphProperties === false) {
      // omit
    } else {
      const epr = opts.endParagraphProperties ?? { lang: "en-US" };
      parts.push(stringifyEndParaRPr(epr, ctx));
    }

    const body = parts.join("");
    return body ? `<a:p>${body}</a:p>` : "<a:p/>";
  },

  parse(el, ctx) {
    const result: ParagraphDescriptorOptions = {};

    // Paragraph properties
    const pPr = findChild(el, "a:pPr");
    if (pPr) {
      result.properties = readParagraphProperties(pPr, ctx) as TextParagraphPropertiesOptions;
    }

    // Collect runs, fields, and breaks in document order.
    const children: (TextRunOptions | TextFieldOptions | TextBreakOptions)[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "a:r") {
        children.push(textRunDesc.parse(child, ctx) as TextRunOptions);
      } else if (child.name === "a:fld") {
        children.push(readTextField(child, ctx));
      } else if (child.name === "a:br") {
        children.push(readBreak(child, ctx));
      }
    }

    const [onlyRun] = children;
    if (
      onlyRun !== undefined &&
      !isTextField(onlyRun) &&
      !isBreak(onlyRun) &&
      children.length === 1 &&
      !result.properties &&
      onlyRun.text !== undefined &&
      Object.keys(onlyRun).length === 1
    ) {
      // Single run with no paragraph properties and only text -> use text shorthand
      result.text = onlyRun.text;
    } else if (children.length > 0) {
      result.children = children;
    }

    // End-paragraph run properties — preserve source value, or mark false so
    // stringify omits it (fresh paragraphs re-emit the default lang marker).
    const endParaRPr = findChild(el, "a:endParaRPr");
    if (endParaRPr) {
      result.endParagraphProperties = runPropertiesDesc.parse(
        endParaRPr,
        ctx,
      ) as TextCharacterPropertiesOptions;
    } else {
      result.endParagraphProperties = false;
    }

    return result as ParagraphDescriptorOptions;
  },
};

// ── Text field (a:fld) + break (a:br) helpers ──

function isTextField(
  child: TextRunOptions | TextFieldOptions | TextBreakOptions,
): child is TextFieldOptions {
  return typeof (child as TextFieldOptions).type === "string";
}

function isBreak(
  child: TextRunOptions | TextFieldOptions | TextBreakOptions,
): child is TextBreakOptions {
  return (child as TextBreakOptions).break === true;
}

function stringifyBreak(opts: TextBreakOptions, ctx: WriteContext): string {
  const rPr = opts.properties ? (runPropertiesDesc.stringify(opts.properties, ctx) ?? "") : "";
  return rPr ? `<a:br>${rPr}</a:br>` : "<a:br/>";
}

function readBreak(el: XmlElement, ctx: ReadContext): TextBreakOptions {
  const result = { break: true } as Mutable<TextBreakOptions>;
  const rPr = findChild(el, "a:rPr");
  if (rPr) result.properties = runPropertiesDesc.parse(rPr, ctx) as TextCharacterPropertiesOptions;
  return result as TextBreakOptions;
}

// a:endParaRPr has the same CT_TextCharacterProperties shape as a:rPr; reuse
// the run-properties serializer under the endParaRPr tag.
function stringifyEndParaRPr(opts: TextCharacterPropertiesOptions, ctx: WriteContext): string {
  return stringifyRunProperties("a:endParaRPr", opts, ctx) || "<a:endParaRPr/>";
}

function stringifyTextField(opts: TextFieldOptions, ctx: WriteContext): string {
  // id is a required GUID on CT_TextField; fall back to a nil UUID so a
  // user-authored field still produces valid OOXML.
  const id = opts.id ?? "{00000000-0000-0000-0000-000000000000}";
  const rPr = opts.properties ? (runPropertiesDesc.stringify(opts.properties, ctx) ?? "") : "";
  // CT_TextField sequence: rPr?, pPr?, t? — a bare <a:pPr/> placeholder (empty
  // options object) still round-trips, so fall back to the empty element.
  const pPrInner = opts.paragraphProperties
    ? (stringifyParagraphProperties(opts.paragraphProperties, ctx) ?? "")
    : "";
  const pPr = pPrInner || (opts.paragraphProperties ? "<a:pPr/>" : "");
  const t = opts.text !== undefined ? `<a:t>${escapeXml(opts.text)}</a:t>` : "";
  return `<a:fld id="${id}" type="${opts.type}">${rPr}${pPr}${t}</a:fld>`;
}

function readTextField(el: XmlElement, ctx: ReadContext): TextFieldOptions {
  const result: TextFieldOptions = { type: String(el.attributes?.["type"] ?? "") };
  const id = el.attributes?.["id"];
  if (id !== undefined) result.id = String(id);
  const rPr = findChild(el, "a:rPr");
  if (rPr) result.properties = runPropertiesDesc.parse(rPr, ctx) as TextCharacterPropertiesOptions;
  const pPr = findChild(el, "a:pPr");
  if (pPr) result.paragraphProperties = readParagraphProperties(pPr, ctx);
  const t = findChild(el, "a:t");
  if (t) {
    result.text = (t.elements ?? [])
      .filter((e) => e.type === "text")
      .map((e) => e.text ?? "")
      .join("");
  }
  return result;
}
