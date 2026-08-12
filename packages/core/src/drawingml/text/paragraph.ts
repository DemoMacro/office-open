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

import { findChild, escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { xsdTextAlign } from "../../util/mappings";
import { textRunDesc } from "./run";
import { runPropertiesDesc } from "./run-properties";
import type { Mutable } from "./run-properties";
import type {
  BulletAutoNumOptions,
  BulletCharOptions,
  BulletOptions,
  BulletPictureOptions,
  BulletStyleOptions,
  ParagraphPropertiesOptions,
  RunOptions,
  RunPropertiesOptions,
  TabStopOptions,
  TextAlignment,
  TextTabAlignment,
  TextFieldOptions,
  BreakOptions,
} from "./types";

/** Descriptor-layer paragraph accumulator (a:p). */
export interface ParagraphDescriptorOptions {
  text?: string;
  properties?: ParagraphPropertiesOptions;
  children?: (RunOptions | TextFieldOptions | BreakOptions | string)[];
  /**
   * End-paragraph run properties (a:endParaRPr). Fresh paragraphs emit a
   * default lang marker; a parsed source preserves its value; false omits it.
   */
  endParagraphProperties?: RunPropertiesOptions | false;
}

// ── Paragraph properties helper (a:pPr) ──

function stringifyParagraphProperties(
  options: ParagraphPropertiesOptions,
  ctx: WriteContext,
): string {
  const children: string[] = [];

  const attrs: string[] = [];
  if (options.alignment) attrs.push(`algn="${xsdTextAlign.to(options.alignment)}"`);
  if (options.indentLevel !== undefined) attrs.push(`lvl="${options.indentLevel}"`);
  if (options.marginIndent !== undefined) attrs.push(`marL="${options.marginIndent}"`);
  if (options.marginRight !== undefined) attrs.push(`marR="${options.marginRight}"`);
  if (options.indent !== undefined) attrs.push(`indent="${options.indent}"`);
  if (options.defTabSize !== undefined) attrs.push(`defTabSz="${options.defTabSize}"`);
  if (options.fontAlignment) attrs.push(`fontAlgn="${options.fontAlignment}"`);
  if (options.rightToLeft !== undefined) attrs.push(`rtl="${options.rightToLeft ? 1 : 0}"`);
  if (options.eastAsianLineBreak !== undefined)
    attrs.push(`eaLnBrk="${options.eastAsianLineBreak ? 1 : 0}"`);
  if (options.latinLineBreak !== undefined)
    attrs.push(`latinLnBrk="${options.latinLineBreak ? 1 : 0}"`);

  // Line spacing
  if (options.lineSpacing !== undefined) {
    children.push(`<a:lnSpc><a:spcPct val="${options.lineSpacing * 100000}"/></a:lnSpc>`);
  }
  if (options.lineSpacingPoints !== undefined) {
    children.push(`<a:lnSpc><a:spcPts val="${options.lineSpacingPoints * 100}"/></a:lnSpc>`);
  }

  // Spacing before/after (XSD CT_TextParagraphProperties order: spcBef, spcAft)
  if (options.marginTop !== undefined) {
    children.push(`<a:spcBef><a:spcPts val="${options.marginTop}"/></a:spcBef>`);
  }
  if (options.marginBottom !== undefined) {
    children.push(`<a:spcAft><a:spcPts val="${options.marginBottom}"/></a:spcAft>`);
  }

  // Bullets
  if (options.bullet) {
    children.push(...stringifyBullet(options.bullet));
  }

  // Tab stops (after bullets, before defRPr)
  if (options.tabStops && options.tabStops.length > 0) {
    children.push(stringifyTabStops(options.tabStops));
  }

  // Default run properties — last child of a:pPr (CT_TextParagraphProperties).
  if (options.defaultRunProperties) {
    const rPr = runPropertiesDesc.stringify(options.defaultRunProperties, ctx) ?? "";
    children.push(rPr.replaceAll("<a:rPr", "<a:defRPr").replaceAll("</a:rPr>", "</a:defRPr>"));
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

  // Color: buClrTx | buClr
  if (options.colorFollowsText) {
    parts.push("<a:buClrTx/>");
  } else if (options.color) {
    parts.push(`<a:buClr><a:srgbClr val="${options.color.replace("#", "")}"/></a:buClr>`);
  }

  // Size: buSzTx | buSzPts | buSzPct
  if (options.sizeFollowsText) {
    parts.push("<a:buSzTx/>");
  } else if (options.sizePoints !== undefined) {
    parts.push(`<a:buSzPts val="${options.sizePoints}"/>`);
  } else if (options.size !== undefined) {
    parts.push(`<a:buSzPct val="${options.size}%"/>`);
  }

  // Font: buFontTx | buFont. Fresh char/autoNum bullets default to Arial.
  if (options.fontFollowsText) {
    parts.push("<a:buFontTx/>");
  } else if (options.font !== undefined) {
    parts.push(
      `<a:buFont typeface="${options.font}" panose="020B0604020202020204" pitchFamily="34" charset="0"/>`,
    );
  } else if (options.type === "char" || options.type === "autoNum") {
    parts.push(
      `<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>`,
    );
  }

  // Bullet type: buChar | buAutoNum | buBlip
  if (options.type === "char") {
    parts.push(`<a:buChar char="${escapeXml(options.char ?? "•")}"/>`);
  } else if (options.type === "autoNum") {
    const buAttrs: string[] = [`type="${options.format ?? "arabicPeriod"}"`];
    if (options.startAt !== undefined) buAttrs.push(`startAt="${options.startAt}"`);
    parts.push(`<a:buAutoNum ${buAttrs.join(" ")}/>`);
  } else if (options.type === "picture") {
    parts.push(`<a:buBlip><a:blip r:embed="${options.embed}"/></a:buBlip>`);
  }

  return parts;
}

function stringifyTabStops(stops: TabStopOptions[]): string {
  const tabs = stops.map((t) => {
    const attrs: string[] = [];
    if (t.position !== undefined) attrs.push(`pos="${t.position}"`);
    if (t.alignment) attrs.push(`algn="${t.alignment}"`);
    return `<a:tab${attrs.length ? " " + attrs.join(" ") : ""}/>`;
  });
  return `<a:tabLst>${tabs.join("")}</a:tabLst>`;
}

function readParagraphProperties(
  el: XmlElement,
  ctx: ReadContext,
): Mutable<ParagraphPropertiesOptions> {
  const result: Mutable<ParagraphPropertiesOptions> = {};

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
      result.fontAlignment = String(
        el.attributes["fontAlgn"],
      ) as ParagraphPropertiesOptions["fontAlignment"];
    if (el.attributes["rtl"] !== undefined)
      result.rightToLeft = String(el.attributes["rtl"]) === "1";
    if (el.attributes["eaLnBrk"] !== undefined)
      result.eastAsianLineBreak = String(el.attributes["eaLnBrk"]) === "1";
    if (el.attributes["latinLnBrk"] !== undefined)
      result.latinLineBreak = String(el.attributes["latinLnBrk"]) === "1";
  }

  // Line spacing
  const lnSpc = findChild(el, "a:lnSpc");
  if (lnSpc) {
    const spcPct = findChild(lnSpc, "a:spcPct");
    if (spcPct?.attributes?.["val"] !== undefined) {
      result.lineSpacing = Number(spcPct.attributes["val"]) / 100000;
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
  } else {
    const buChar = findChild(el, "a:buChar");
    const buAutoNum = findChild(el, "a:buAutoNum");
    const buBlip = findChild(el, "a:buBlip");
    if (buChar || buAutoNum || buBlip) {
      // Shared color/size/font style — each dimension is a choice.
      const style: Mutable<BulletStyleOptions> = {};
      if (findChild(el, "a:buClrTx")) {
        style.colorFollowsText = true;
      } else {
        const buClr = findChild(el, "a:buClr");
        if (buClr) {
          const srgb = findChild(buClr, "a:srgbClr");
          if (srgb?.attributes?.["val"]) style.color = String(srgb.attributes["val"]);
        }
      }
      if (findChild(el, "a:buSzTx")) {
        style.sizeFollowsText = true;
      } else {
        const buSzPts = findChild(el, "a:buSzPts");
        if (buSzPts?.attributes?.["val"]) {
          style.sizePoints = Number(buSzPts.attributes["val"]);
        } else {
          const buSzPct = findChild(el, "a:buSzPct");
          if (buSzPct?.attributes?.["val"])
            style.size = Number(String(buSzPct.attributes["val"]).replace("%", ""));
        }
      }
      if (findChild(el, "a:buFontTx")) {
        style.fontFollowsText = true;
      } else {
        const buFont = findChild(el, "a:buFont");
        if (buFont?.attributes?.["typeface"]) style.font = String(buFont.attributes["typeface"]);
      }

      if (buBlip) {
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
        if (buAutoNum!.attributes?.["type"]) bullet.format = String(buAutoNum!.attributes["type"]);
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
          tab.alignment = child.attributes["algn"] as TextTabAlignment;
        tabs.push(tab);
      }
    }
    if (tabs.length > 0) result.tabStops = tabs;
  }

  // Default run properties — last child of a:pPr (CT_TextParagraphProperties).
  const defRPr = findChild(el, "a:defRPr");
  if (defRPr) {
    result.defaultRunProperties = runPropertiesDesc.parse(defRPr, ctx) as RunPropertiesOptions;
  }

  return result as ParagraphPropertiesOptions;
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
      parts.push(textRunDesc.stringify({ text: opts.text }, ctx) ?? "");
    }

    // Children (text runs + fields + line breaks; a bare string is shorthand
    // for a single text-only run).
    if (opts.children) {
      for (const child of opts.children) {
        if (typeof child === "string") {
          parts.push(textRunDesc.stringify({ text: child }, ctx) ?? "");
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
      result.properties = readParagraphProperties(pPr, ctx) as ParagraphPropertiesOptions;
    }

    // Collect runs, fields, and breaks in document order.
    const children: (RunOptions | TextFieldOptions | BreakOptions)[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "a:r") {
        children.push(textRunDesc.parse(child, ctx) as RunOptions);
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
      ) as RunPropertiesOptions;
    } else {
      result.endParagraphProperties = false;
    }

    return result as ParagraphDescriptorOptions;
  },
};

// ── Text field (a:fld) + break (a:br) helpers ──

function isTextField(
  child: RunOptions | TextFieldOptions | BreakOptions,
): child is TextFieldOptions {
  return typeof (child as TextFieldOptions).type === "string";
}

function isBreak(child: RunOptions | TextFieldOptions | BreakOptions): child is BreakOptions {
  return (child as BreakOptions).break === true;
}

function stringifyBreak(opts: BreakOptions, ctx: WriteContext): string {
  const rPr = opts.properties ? (runPropertiesDesc.stringify(opts.properties, ctx) ?? "") : "";
  return rPr ? `<a:br>${rPr}</a:br>` : "<a:br/>";
}

function readBreak(el: XmlElement, ctx: ReadContext): BreakOptions {
  const result = { break: true } as Mutable<BreakOptions>;
  const rPr = findChild(el, "a:rPr");
  if (rPr) result.properties = runPropertiesDesc.parse(rPr, ctx) as RunPropertiesOptions;
  return result as BreakOptions;
}

// a:endParaRPr has the same CT_TextCharacterProperties shape as a:rPr; reuse
// runPropertiesDesc and relabel the emitted tag.
function stringifyEndParaRPr(opts: RunPropertiesOptions, ctx: WriteContext): string {
  const rPr = runPropertiesDesc.stringify(opts, ctx) ?? "";
  if (!rPr) return "<a:endParaRPr/>";
  return rPr.replaceAll("<a:rPr", "<a:endParaRPr").replaceAll("</a:rPr>", "</a:endParaRPr>");
}

function stringifyTextField(opts: TextFieldOptions, ctx: WriteContext): string {
  // id is a required GUID on CT_TextField; fall back to a nil UUID so a
  // user-authored field still produces valid OOXML.
  const id = opts.id ?? "{00000000-0000-0000-0000-000000000000}";
  const rPr = opts.properties ? (runPropertiesDesc.stringify(opts.properties, ctx) ?? "") : "";
  const t = opts.text !== undefined ? `<a:t>${escapeXml(opts.text)}</a:t>` : "";
  return `<a:fld id="${id}" type="${opts.type}">${rPr}${t}</a:fld>`;
}

function readTextField(el: XmlElement, ctx: ReadContext): TextFieldOptions {
  const result: TextFieldOptions = { type: String(el.attributes?.["type"] ?? "") };
  const id = el.attributes?.["id"];
  if (id !== undefined) result.id = String(id);
  const rPr = findChild(el, "a:rPr");
  if (rPr) result.properties = runPropertiesDesc.parse(rPr, ctx) as RunPropertiesOptions;
  const t = findChild(el, "a:t");
  if (t) {
    result.text = (t.elements ?? [])
      .filter((e) => e.type === "text")
      .map((e) => e.text ?? "")
      .join("");
  }
  return result;
}
