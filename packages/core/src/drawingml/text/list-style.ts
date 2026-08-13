/**
 * Text list style (CT_TextListStyle) descriptor.
 *
 * CT_TextListStyle is the deeply-nested structure behind p:defaultTextStyle
 * (presentation) and p:txStyles titleStyle/bodyStyle/otherStyle (slide master):
 * up to 9 outline levels (lvl1pPr..lvl9pPr), each a CT_TextParagraphProperties
 * carrying alignment, indent, spacing, bullet, and a default-run (defRPr) block.
 *
 * Stringify models MS Office's byte layout exactly so the structured default
 * reproduces the prior verbatim block bit-for-bit. Parse is the inverse.
 *
 * @module
 */

import type { Element } from "@office-open/xml";
import { escapeXml, findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";

// ── Types ──

export interface TextListStyleRunOptions {
  /** Font size in points (e.g., 44 = 44pt). */
  size?: number;
  /** Character kerning in points (e.g., 12 = 12pt). */
  kern?: number;
  /** Solid fill as a theme color token (e.g. "tx1"). */
  schemeColor?: string;
  /** Latin script typeface (e.g. "+mj-lt"). */
  latin?: string;
  /** East-Asian script typeface (e.g. "+mj-ea"). */
  ea?: string;
  /** Complex-script typeface (e.g. "+mj-cs"). */
  cs?: string;
}

export interface TextListStyleBulletOptions {
  type: "none" | "char";
  /** Bullet glyph (type="char"); defaults to "•". */
  char?: string;
  /** Bullet font typeface (type="char"); defaults to "Arial". */
  font?: string;
}

export interface TextListStyleLevelOptions {
  alignment?: string;
  marginIndent?: number;
  indent?: number;
  defaultTabSize?: number;
  rtl?: boolean;
  /** East-Asian line breaking (default true in MS Office). */
  eastAsianLineBreak?: boolean;
  /** Latin line breaking (default false). */
  latinLineBreak?: boolean;
  hangingPunctuation?: boolean;
  /** Line spacing as a percentage (e.g., 90 = 90%). */
  lineSpacingPercent?: number;
  /** Space before as a percentage (e.g., 0 = 0%). */
  spaceBeforePercent?: number;
  /** Space before in points (e.g., 5 = 5pt). */
  spaceBeforePoints?: number;
  bullet?: TextListStyleBulletOptions;
  defaultRun?: TextListStyleRunOptions;
}

/** A title/body/other group: an optional empty defPPr plus up to 9 levels. */
export interface TextListStyleGroupOptions {
  /** Emit an empty <a:defPPr><a:defRPr/></a:defPPr> (otherStyle carries one). */
  emptyDefaultParagraph?: boolean;
  /** Levels 1-9; index 0 = lvl1pPr. `undefined` omits the level. */
  levels?: (TextListStyleLevelOptions | undefined)[];
}

export interface TextListStyleOptions {
  title?: TextListStyleGroupOptions;
  body?: TextListStyleGroupOptions;
  other?: TextListStyleGroupOptions;
}

// ── Stringify ──

function stringifyRun(run: TextListStyleRunOptions | undefined): string {
  if (!run) return "";
  const attrs: string[] = [];
  if (run.size !== undefined) attrs.push(`sz="${Math.round(run.size * 100)}"`);
  if (run.kern !== undefined) attrs.push(`kern="${Math.round(run.kern * 100)}"`);
  const fillXml = run.schemeColor
    ? `<a:solidFill><a:schemeClr val="${run.schemeColor}"/></a:solidFill>`
    : "";
  const latinXml = run.latin ? `<a:latin typeface="${run.latin}"/>` : "";
  const eaXml = run.ea ? `<a:ea typeface="${run.ea}"/>` : "";
  const csXml = run.cs ? `<a:cs typeface="${run.cs}"/>` : "";
  const inner = `${fillXml}${latinXml}${eaXml}${csXml}`;
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  return `<a:defRPr${attrStr}>${inner}</a:defRPr>`;
}

function stringifyBullet(b: TextListStyleBulletOptions): string {
  if (b.type === "none") return "<a:buNone/>";
  const typeface = b.font ?? "Arial";
  const char = escapeXml(b.char ?? "•");
  return `<a:buFont typeface="${typeface}" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="${char}"/>`;
}

/** Emit one <a:lvlNpPr>. Byte layout matches MS Office's txStyles output. */
function stringifyLevel(level: number, opts: TextListStyleLevelOptions): string {
  const attrs: string[] = [];
  // Attribute order: marL, indent, algn, defTabSz, rtl, eaLnBrk, latinLnBrk, hangingPunct
  if (opts.marginIndent !== undefined) attrs.push(`marL="${opts.marginIndent}"`);
  if (opts.indent !== undefined) attrs.push(`indent="${opts.indent}"`);
  if (opts.alignment) attrs.push(`algn="${opts.alignment}"`);
  if (opts.defaultTabSize !== undefined) attrs.push(`defTabSz="${opts.defaultTabSize}"`);
  if (opts.rtl !== undefined) attrs.push(`rtl="${opts.rtl ? 1 : 0}"`);
  if (opts.eastAsianLineBreak !== undefined)
    attrs.push(`eaLnBrk="${opts.eastAsianLineBreak ? 1 : 0}"`);
  if (opts.latinLineBreak !== undefined) attrs.push(`latinLnBrk="${opts.latinLineBreak ? 1 : 0}"`);
  if (opts.hangingPunctuation !== undefined)
    attrs.push(`hangingPunct="${opts.hangingPunctuation ? 1 : 0}"`);

  const children: string[] = [];
  // Child order: lnSpc, spcBef, bullet, defRPr
  if (opts.lineSpacingPercent !== undefined)
    children.push(
      `<a:lnSpc><a:spcPct val="${Math.round(opts.lineSpacingPercent * 1000)}"/></a:lnSpc>`,
    );
  if (opts.spaceBeforePercent !== undefined)
    children.push(
      `<a:spcBef><a:spcPct val="${Math.round(opts.spaceBeforePercent * 1000)}"/></a:spcBef>`,
    );
  else if (opts.spaceBeforePoints !== undefined)
    children.push(
      `<a:spcBef><a:spcPts val="${Math.round(opts.spaceBeforePoints * 100)}"/></a:spcBef>`,
    );
  if (opts.bullet) children.push(stringifyBullet(opts.bullet));
  const runXml = stringifyRun(opts.defaultRun);
  if (runXml) children.push(runXml);

  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  if (children.length === 0) return `<a:lvl${level}pPr${attrStr}/>`;
  return `<a:lvl${level}pPr${attrStr}>${children.join("")}</a:lvl${level}pPr>`;
}

function stringifyGroup(tag: string, group: TextListStyleGroupOptions | undefined): string {
  if (!group) return "";
  const parts: string[] = [];
  if (group.emptyDefaultParagraph) parts.push("<a:defPPr><a:defRPr/></a:defPPr>");
  const levels = group.levels ?? [];
  for (let i = 0; i < Math.min(levels.length, 9); i++) {
    const lvl = levels[i];
    if (lvl) parts.push(stringifyLevel(i + 1, lvl));
  }
  return `<${tag}>${parts.join("")}</${tag}>`;
}

/**
 * Stringify the three text-style groups (CT_SlideMasterTextStyles). Group tags
 * use the p: prefix (PML); level tags inside are a: (DrawingML). Caller wraps
 * with `<p:txStyles>…</p:txStyles>`.
 */
export function stringifyTextListStyle(opts: TextListStyleOptions): string {
  return `${stringifyGroup("p:titleStyle", opts.title)}${stringifyGroup("p:bodyStyle", opts.body)}${stringifyGroup("p:otherStyle", opts.other)}`;
}

// ── Parse ──

function parseRun(el: Element | undefined): TextListStyleRunOptions | undefined {
  if (!el) return undefined;
  const run: TextListStyleRunOptions = {};
  if (el.attributes) {
    if (el.attributes["sz"] !== undefined) run.size = Number(el.attributes["sz"]) / 100;
    if (el.attributes["kern"] !== undefined) run.kern = Number(el.attributes["kern"]) / 100;
  }
  const solidFill = findChild(el, "a:solidFill");
  if (solidFill) {
    const schemeClr = findChild(solidFill, "a:schemeClr");
    if (schemeClr?.attributes?.["val"]) run.schemeColor = String(schemeClr.attributes["val"]);
  }
  const latin = findChild(el, "a:latin");
  if (latin?.attributes?.["typeface"]) run.latin = String(latin.attributes["typeface"]);
  const ea = findChild(el, "a:ea");
  if (ea?.attributes?.["typeface"]) run.ea = String(ea.attributes["typeface"]);
  const cs = findChild(el, "a:cs");
  if (cs?.attributes?.["typeface"]) run.cs = String(cs.attributes["typeface"]);
  return Object.keys(run).length > 0 ? run : undefined;
}

function parseLevel(el: Element): TextListStyleLevelOptions {
  const lvl: TextListStyleLevelOptions = {};
  // nativeTypeAttributes (opc parser) coerces "1"/"0" to numbers, so a strict
  // `=== "1"` check silently fails; normalize via String() before comparing.
  const isOn = (raw: unknown): boolean => String(raw) === "1";
  if (el.attributes) {
    const a = el.attributes;
    if (a["algn"] !== undefined) lvl.alignment = String(a["algn"]);
    if (a["marL"] !== undefined) lvl.marginIndent = Number(a["marL"]);
    if (a["indent"] !== undefined) lvl.indent = Number(a["indent"]);
    if (a["defTabSz"] !== undefined) lvl.defaultTabSize = Number(a["defTabSz"]);
    if (a["rtl"] !== undefined) lvl.rtl = isOn(a["rtl"]);
    if (a["eaLnBrk"] !== undefined) lvl.eastAsianLineBreak = isOn(a["eaLnBrk"]);
    if (a["latinLnBrk"] !== undefined) lvl.latinLineBreak = isOn(a["latinLnBrk"]);
    if (a["hangingPunct"] !== undefined) lvl.hangingPunctuation = isOn(a["hangingPunct"]);
  }
  const lnSpc = findChild(el, "a:lnSpc");
  if (lnSpc) {
    const spcPct = findChild(lnSpc, "a:spcPct");
    if (spcPct?.attributes?.["val"] !== undefined)
      lvl.lineSpacingPercent = Number(spcPct.attributes["val"]) / 1000;
  }
  const spcBef = findChild(el, "a:spcBef");
  if (spcBef) {
    const spcPct = findChild(spcBef, "a:spcPct");
    if (spcPct?.attributes?.["val"] !== undefined)
      lvl.spaceBeforePercent = Number(spcPct.attributes["val"]) / 1000;
    const spcPts = findChild(spcBef, "a:spcPts");
    if (spcPts?.attributes?.["val"] !== undefined)
      lvl.spaceBeforePoints = Number(spcPts.attributes["val"]) / 100;
  }
  if (findChild(el, "a:buNone")) {
    lvl.bullet = { type: "none" };
  } else {
    const buChar = findChild(el, "a:buChar");
    if (buChar) {
      const buFont = findChild(el, "a:buFont");
      lvl.bullet = {
        type: "char",
        char: buChar.attributes?.["char"] ? String(buChar.attributes["char"]) : undefined,
        font: buFont?.attributes?.["typeface"] ? String(buFont.attributes["typeface"]) : undefined,
      };
    }
  }
  const defRPr = findChild(el, "a:defRPr");
  const run = parseRun(defRPr);
  if (run) lvl.defaultRun = run;
  return lvl;
}

function parseGroup(el: Element | undefined): TextListStyleGroupOptions | undefined {
  if (!el) return undefined;
  const group: TextListStyleGroupOptions = { levels: [] };
  const defPPr = findChild(el, "a:defPPr");
  if (defPPr) group.emptyDefaultParagraph = true;
  for (let i = 1; i <= 9; i++) {
    const lvlEl = findChild(el, `a:lvl${i}pPr`);
    group.levels!.push(lvlEl ? parseLevel(lvlEl) : undefined);
  }
  const hasContent = group.emptyDefaultParagraph || group.levels!.some((l) => l !== undefined);
  return hasContent ? group : undefined;
}

/** Parse CT_SlideMasterTextStyles (the three p:titleStyle/p:bodyStyle/p:otherStyle groups). */
export function parseTextListStyle(el: Element): TextListStyleOptions {
  return {
    title: parseGroup(findChild(el, "p:titleStyle")),
    body: parseGroup(findChild(el, "p:bodyStyle")),
    other: parseGroup(findChild(el, "p:otherStyle")),
  };
}

// ── Descriptor ──

export const textListStyleDesc: CustomDescriptor<TextListStyleOptions> = {
  kind: "custom",
  stringify(opts) {
    return stringifyTextListStyle(opts);
  },
  parse(el) {
    return parseTextListStyle(el);
  },
};

// ── MS Office standard master text styles (structured form) ──

const MJ_RUN = { schemeColor: "tx1", latin: "+mj-lt", ea: "+mj-ea", cs: "+mj-cs" } as const;
const MN_RUN = { schemeColor: "tx1", latin: "+mn-lt", ea: "+mn-ea", cs: "+mn-cs" } as const;
const BASE_ATTRS = {
  alignment: "l",
  defaultTabSize: 914400,
  rtl: false,
  eastAsianLineBreak: true,
  latinLineBreak: false,
  hangingPunctuation: true,
} as const;

const BODY_MARL = [228600, 685800, 1143000, 1600200, 2057400, 2514600, 2971800, 3429000, 3886200];
const BODY_SZ = [28, 24, 20, 18, 18, 18, 18, 18, 18];
const OTHER_MARL = [0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600];

/** MS Office's default p:txStyles — title/body/other, 9 outline levels. */
export const DEFAULT_TEXT_LIST_STYLE: TextListStyleOptions = {
  title: {
    levels: [
      {
        ...BASE_ATTRS,
        lineSpacingPercent: 90,
        spaceBeforePercent: 0,
        bullet: { type: "none" },
        defaultRun: { size: 44, kern: 12, ...MJ_RUN },
      },
    ],
  },
  body: {
    levels: BODY_MARL.map((marL, i) => ({
      ...BASE_ATTRS,
      marginIndent: marL,
      indent: -228600,
      lineSpacingPercent: 90,
      spaceBeforePoints: i === 0 ? undefined : 5,
      spaceBeforePercent: i === 0 ? 0 : undefined,
      bullet: { type: "char", char: "•", font: "Arial" },
      defaultRun: { size: BODY_SZ[i]!, kern: 12, ...MN_RUN },
    })),
  },
  other: {
    emptyDefaultParagraph: true,
    levels: OTHER_MARL.map((marL, i) => ({
      ...BASE_ATTRS,
      marginIndent: marL,
      indent: i === 8 ? -228600 : undefined,
      defaultRun: { size: 18, kern: 12, ...MN_RUN },
    })),
  },
};
