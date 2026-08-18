/**
 * Text list style descriptors.
 *
 * CT_TextListStyle is the level list behind a:lstStyle (txBody), p:notesStyle
 * and p:defaultTextStyle: an optional defPPr plus up to 9 outline levels
 * (lvl1pPr..lvl9pPr), each a CT_TextParagraphProperties reusing the shared
 * paragraph model. CT_SlideMasterTextStyles groups three of those lists as
 * p:titleStyle/p:bodyStyle/p:otherStyle on the slide master.
 *
 * Stringify models MS Office's byte layout exactly (XSD declaration order)
 * so the structured defaults reproduce Office-authored bytes bit-for-bit.
 * Parse is the inverse via the shared paragraph reader.
 *
 * @module
 */

import type { Element } from "@office-open/xml";
import { findChild, stringify } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { readParagraphProperties, stringifyParagraphPropertiesElement } from "./paragraph";
import type { ParagraphPropertiesOptions } from "./types";

// ── Types ──

/**
 * A bare CT_TextListStyle — a:defPPr plus up to 9 outline levels. The shape
 * of a:lstStyle inside a txBody, p:notesStyle, and p:defaultTextStyle.
 */
export interface TextListStyleOptions {
  /** a:defPPr — defaults applied before any level. */
  defaultParagraph?: ParagraphPropertiesOptions;
  /** Levels 1-9; index 0 = lvl1pPr. `null` omits the level (JSON form of
   *  an undefined array slot — the position still pins the level number). */
  levels?: (ParagraphPropertiesOptions | null)[];
  /** Trailing a:extLst verbatim inner XML; `""` preserves a bare <a:extLst/>. */
  ext?: string;
}

/** CT_SlideMasterTextStyles — the master's title/body/other style groups. */
export interface TextStylesOptions {
  title?: TextListStyleOptions;
  body?: TextListStyleOptions;
  other?: TextListStyleOptions;
}

// ── Stringify ──

function stringifyGroup(
  tag: string,
  group: TextListStyleOptions | undefined,
  ctx: WriteContext,
): string {
  if (!group) return "";
  const parts: string[] = [];
  if (group.defaultParagraph) {
    parts.push(stringifyParagraphPropertiesElement("a:defPPr", group.defaultParagraph, ctx));
  }
  const levels = group.levels ?? [];
  for (let i = 0; i < Math.min(levels.length, 9); i++) {
    const lvl = levels[i];
    if (lvl) parts.push(stringifyParagraphPropertiesElement(`a:lvl${i + 1}pPr`, lvl, ctx));
  }
  if (group.ext !== undefined) parts.push(`<a:extLst>${group.ext}</a:extLst>`);
  return `<${tag}>${parts.join("")}</${tag}>`;
}

/**
 * Stringify the three master text-style groups (CT_SlideMasterTextStyles).
 * Group tags use the p: prefix (PML); level tags inside are a: (DrawingML).
 * Caller wraps with `<p:txStyles>…</p:txStyles>`.
 */
export function stringifyTextStyles(opts: TextStylesOptions, ctx: WriteContext): string {
  return (
    `${stringifyGroup("p:titleStyle", opts.title, ctx)}` +
    `${stringifyGroup("p:bodyStyle", opts.body, ctx)}` +
    `${stringifyGroup("p:otherStyle", opts.other, ctx)}`
  );
}

/** Emit a bare CT_TextListStyle under a caller-chosen tag, e.g. p:notesStyle. */
export function stringifyTextListStyleTag(
  tag: string,
  group: TextListStyleOptions | undefined,
  ctx: WriteContext,
): string {
  return stringifyGroup(tag, group, ctx);
}

// ── Parse ──

function parseGroup(el: Element | undefined, ctx: ReadContext): TextListStyleOptions | undefined {
  if (!el) return undefined;
  const group: TextListStyleOptions = {};
  const defPPr = findChild(el, "a:defPPr");
  if (defPPr) group.defaultParagraph = readParagraphProperties(defPPr, ctx);
  const levels: (ParagraphPropertiesOptions | null)[] = [];
  for (let i = 1; i <= 9; i++) {
    const lvlEl = findChild(el, `a:lvl${i}pPr`);
    levels.push(lvlEl ? readParagraphProperties(lvlEl, ctx) : null);
  }
  if (group.defaultParagraph || levels.some((l) => l !== null)) group.levels = levels;
  const extLst = findChild(el, "a:extLst");
  if (extLst) group.ext = stringify(extLst);
  if (group.levels || group.ext !== undefined) return group;
  return undefined;
}

/** Parse CT_SlideMasterTextStyles (the three p:titleStyle/p:bodyStyle/p:otherStyle groups). */
export function parseTextStyles(el: Element, ctx: ReadContext): TextStylesOptions {
  return {
    title: parseGroup(findChild(el, "p:titleStyle"), ctx),
    body: parseGroup(findChild(el, "p:bodyStyle"), ctx),
    other: parseGroup(findChild(el, "p:otherStyle"), ctx),
  };
}

// ── Descriptors ──

/** A bare CT_TextListStyle (a:lstStyle in a txBody). */
export const textListStyleDesc: CustomDescriptor<TextListStyleOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];
    if (opts.defaultParagraph) {
      parts.push(stringifyParagraphPropertiesElement("a:defPPr", opts.defaultParagraph, ctx));
    }
    const levels = opts.levels ?? [];
    for (let i = 0; i < Math.min(levels.length, 9); i++) {
      const lvl = levels[i];
      if (lvl) parts.push(stringifyParagraphPropertiesElement(`a:lvl${i + 1}pPr`, lvl, ctx));
    }
    if (opts.ext !== undefined) parts.push(`<a:extLst>${opts.ext}</a:extLst>`);
    return parts.join("");
  },
  parse(el, ctx) {
    return parseGroup(el, ctx) ?? {};
  },
};

/** CT_SlideMasterTextStyles descriptor (p:titleStyle/p:bodyStyle/p:otherStyle). */
export const textStylesDesc: CustomDescriptor<TextStylesOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    return stringifyTextStyles(opts, ctx);
  },
  parse(el, ctx) {
    return parseTextStyles(el, ctx);
  },
};

// ── MS Office standard master text styles (structured form) ──

const MJ_RUN = {
  fill: { type: "solid", color: { value: "tx1" } },
  font: { latin: "+mj-lt", eastAsia: "+mj-ea", complexScript: "+mj-cs" },
} as const;
const MN_RUN = {
  fill: { type: "solid", color: { value: "tx1" } },
  font: { latin: "+mn-lt", eastAsia: "+mn-ea", complexScript: "+mn-cs" },
} as const;
const BASE_ATTRS = {
  alignment: "left",
  defTabSize: 914400,
  rightToLeft: false,
  eastAsianLineBreak: true,
  latinLineBreak: false,
  hangingPunctuation: true,
} as const;

const BODY_MARL = [228600, 685800, 1143000, 1600200, 2057400, 2514600, 2971800, 3429000, 3886200];
const BODY_SZ = [28, 24, 20, 18, 18, 18, 18, 18, 18];
const OTHER_MARL = [0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600];

/** MS Office's default p:txStyles — title/body/other, 9 outline levels. */
export const DEFAULT_TEXT_STYLES: TextStylesOptions = {
  title: {
    levels: [
      {
        ...BASE_ATTRS,
        lineSpacingPercent: 90,
        spaceBeforePercent: 0,
        bullet: { type: "none" },
        defaultRunProperties: { size: 44, kern: 12, ...MJ_RUN },
      },
    ],
  },
  body: {
    levels: BODY_MARL.map((marL, i) => ({
      ...BASE_ATTRS,
      marginIndent: marL,
      indent: -228600,
      lineSpacingPercent: 90,
      spaceBeforePercent: i === 0 ? 0 : undefined,
      spaceBefore: i === 0 ? undefined : 5,
      bullet: { type: "char", char: "•", font: "Arial" },
      defaultRunProperties: { size: BODY_SZ[i]!, kern: 12, ...MN_RUN },
    })),
  },
  other: {
    defaultParagraph: { defaultRunProperties: {} },
    levels: OTHER_MARL.map((marL, i) => ({
      ...BASE_ATTRS,
      marginIndent: marL,
      indent: i === 8 ? -228600 : undefined,
      defaultRunProperties: { size: 18, kern: 12, ...MN_RUN },
    })),
  },
};
