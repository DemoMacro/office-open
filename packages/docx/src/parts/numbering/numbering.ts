import { uniqueNumericIdCreator, convertInchesToTwip } from "@office-open/core";
import { decimalNumber } from "@office-open/core";
/**
 * Numbering module for WordprocessingML documents.
 *
 * Numbering provides support for numbered and bulleted lists.
 * Pure string serialization — no XmlComponent inheritance.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * @module
 */
import { attr, attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { AlignmentType } from "@parts/paragraph";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";

import type { DocxReadContext } from "../../context";
import { stringifyParagraphProperties, stringifyRunProperties } from "../paragraph/stringify";
import { LevelFormat } from "./level";
import type { LevelsOptions } from "./level";

/**
 * Options for configuring numbering definitions.
 */
export interface NumberingOptions {
  /** Array of numbering configurations, each with levels and a reference name. */
  config: {
    levels: LevelsOptions[];
    reference: string;
    extraOptions?: AbstractNumberingExtraOptions;
  }[];
  /** Numbering cleanup ID (w:numIdMacAtCleanup) */
  numIdMacAtCleanup?: number;
  /** Picture bullet definitions for numbering (w:numPicBullet) */
  numPicBullets?: {
    numPicBulletId: number;
    pict?: string;
  }[];
}

/** Namespace attributes for w:numbering (pre-computed constant). */
const NUMBERING_ATTRS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'mc:Ignorable="w14 w15 wp14"';

/** Default bullet levels (9 levels: 0-8). */
const DEFAULT_BULLET_LEVELS: LevelsOptions[] = [
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 0,
    style: {
      paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: convertInchesToTwip(0.5) } },
    },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 1,
    style: {
      paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: convertInchesToTwip(1) } },
    },
    text: "○",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 2,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 2160 } } },
    text: "■",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 3,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 2880 } } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 4,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 3600 } } },
    text: "○",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 5,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 4320 } } },
    text: "■",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 6,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 5040 } } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 7,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 5760 } } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 8,
    style: { paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 6480 } } },
    text: "●",
  },
];

/**
 * Numbering definitions for a WordprocessingML document.
 *
 * Pure data accumulator — no XmlComponent inheritance.
 * Serializes via `serialize()` producing a complete XML part.
 */
export class Numbering {
  private abstractNumberingData = new Map<
    string,
    { id: number; levels: LevelsOptions[]; extraOptions?: AbstractNumberingExtraOptions }
  >();
  private concreteNumberingData = new Map<
    string,
    {
      numId: number;
      abstractNumId: number;
      reference: string;
      instance: number;
      overrideLevels?: { num: number; start?: number }[];
    }
  >();
  private referenceConfigMap = new Map<string, LevelsOptions[]>();
  private abstractNumUniqueNumericId = uniqueNumericIdCreator();
  private concreteNumUniqueNumericId = uniqueNumericIdCreator(1);
  private _numIdMacAtCleanup?: number;
  private _numPicBullets?: { numPicBulletId: number; pict?: string }[];

  public constructor(options: NumberingOptions) {
    this._numIdMacAtCleanup = options.numIdMacAtCleanup;
    this._numPicBullets = options.numPicBullets;

    // Only inject the default bullet numbering when the caller supplied no
    // numbering definitions. Round-tripped documents carry their own, so
    // injecting a default would inflate the part (extra abstractNum + 9 levels).
    if (options.config.length === 0) {
      const defaultAbstractId = this.abstractNumUniqueNumericId();
      this.abstractNumberingData.set("default-bullet-numbering", {
        id: defaultAbstractId,
        levels: DEFAULT_BULLET_LEVELS,
      });
      this.concreteNumberingData.set("default-bullet-numbering", {
        abstractNumId: defaultAbstractId,
        instance: 0,
        numId: 1,
        overrideLevels: [{ num: 0, start: 1 }],
        reference: "default-bullet-numbering",
      });
    }

    for (const con of options.config) {
      this.abstractNumberingData.set(con.reference, {
        id: this.abstractNumUniqueNumericId(),
        levels: con.levels,
        extraOptions: con.extraOptions,
      });
      this.referenceConfigMap.set(con.reference, con.levels);
    }
  }

  /** Serialize to word/numbering.xml content (with XML declaration). */
  public serialize(): string {
    const parts: string[] = [];
    parts.push(`<w:numbering ${NUMBERING_ATTRS}>`);

    // numPicBullet elements come first (XSD order)
    if (this._numPicBullets) {
      for (const bullet of this._numPicBullets) {
        if (bullet.pict) {
          parts.push(
            `<w:numPicBullet w:numPicBulletId="${bullet.numPicBulletId}">${bullet.pict}</w:numPicBullet>`,
          );
        } else {
          parts.push(`<w:numPicBullet w:numPicBulletId="${bullet.numPicBulletId}"/>`);
        }
      }
    }

    for (const an of this.abstractNumberingData.values()) {
      parts.push(stringifyAbstractNumbering(an.id, an.levels, an.extraOptions));
    }
    for (const cn of this.concreteNumberingData.values()) {
      parts.push(stringifyConcreteNumbering(cn));
    }
    if (this._numIdMacAtCleanup !== undefined) {
      parts.push(`<w:numIdMacAtCleanup w:val="${decimalNumber(this._numIdMacAtCleanup)}"/>`);
    }

    parts.push("</w:numbering>");
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + parts.join("");
  }

  /**
   * Creates a concrete numbering instance from an abstract numbering definition.
   */
  public createConcreteNumberingInstance(reference: string, instance: number): void {
    const abstractNumbering = this.abstractNumberingData.get(reference);
    if (!abstractNumbering) return;

    const fullReference = `${reference}-${instance}`;
    if (this.concreteNumberingData.has(fullReference)) return;

    const referenceConfigLevels = this.referenceConfigMap.get(reference);
    const firstLevelStartNumber = referenceConfigLevels?.[0]?.start;

    // Only emit a lvlOverride when the instance's start differs from the
    // abstract default (1) — otherwise it is redundant and inflates the part.
    const overrideLevels =
      typeof firstLevelStartNumber === "number" &&
      Number.isInteger(firstLevelStartNumber) &&
      firstLevelStartNumber !== 1
        ? [{ num: 0, start: firstLevelStartNumber }]
        : undefined;

    this.concreteNumberingData.set(fullReference, {
      abstractNumId: abstractNumbering.id,
      instance,
      numId: this.concreteNumUniqueNumericId(),
      overrideLevels,
      reference,
    });
  }

  /** Gets all concrete numbering instances. */
  public get concreteNumbering(): { numId: number; reference: string; instance: number }[] {
    return [...this.concreteNumberingData.values()];
  }

  /** Gets all reference configurations. */
  public get referenceConfig(): LevelsOptions[][] {
    return [...this.referenceConfigMap.values()];
  }
}

// ── Types ──

/** Extra options for abstract numbering. Re-exported from abstract-numbering module. */
interface AbstractNumberingExtraOptions {
  nsid?: string;
  /** w:multiLevelType value (singleLevel/multilevel/hybridMultilevel). */
  multiLevelType?: string;
  /** w15:restartNumberingAfterBreak attribute on w:abstractNum. Omitted when undefined. */
  restartNumberingAfterBreak?: boolean;
  tmpl?: string;
  name?: string;
  styleLink?: string;
  numStyleLink?: string;
}

// ── Pure function serializers ──

function stringifyAbstractNumbering(
  id: number,
  levels: LevelsOptions[],
  extraOptions?: AbstractNumberingExtraOptions,
): string {
  const parts: string[] = [];
  // w15:restartNumberingAfterBreak is optional (w15 extension); only emit when
  // explicitly carried so round-trip matches sources that omit it.
  const restartAttr =
    extraOptions?.restartNumberingAfterBreak !== undefined
      ? ` w15:restartNumberingAfterBreak="${extraOptions.restartNumberingAfterBreak ? 1 : 0}"`
      : "";
  parts.push(`<w:abstractNum w:abstractNumId="${decimalNumber(id)}"${restartAttr}>`);

  if (extraOptions?.nsid !== undefined) {
    parts.push(`<w:nsid w:val="${extraOptions.nsid}"/>`);
  }
  parts.push(`<w:multiLevelType w:val="${extraOptions?.multiLevelType ?? "hybridMultilevel"}"/>`);
  if (extraOptions?.tmpl !== undefined) {
    parts.push(`<w:tmpl w:val="${extraOptions.tmpl}"/>`);
  }
  if (extraOptions?.name !== undefined) {
    parts.push(`<w:name w:val="${extraOptions.name}"/>`);
  }
  if (extraOptions?.styleLink !== undefined) {
    parts.push(`<w:styleLink w:val="${extraOptions.styleLink}"/>`);
  }
  if (extraOptions?.numStyleLink !== undefined) {
    parts.push(`<w:numStyleLink w:val="${extraOptions.numStyleLink}"/>`);
  }

  for (const level of levels) {
    parts.push(stringifyLevel(level));
  }

  parts.push("</w:abstractNum>");
  return parts.join("");
}

function stringifyConcreteNumbering(cn: {
  numId: number;
  abstractNumId: number;
  overrideLevels?: { num: number; start?: number }[];
}): string {
  const parts: string[] = [];
  parts.push(`<w:num w:numId="${decimalNumber(cn.numId)}">`);
  parts.push(`<w:abstractNumId w:val="${decimalNumber(cn.abstractNumId)}"/>`);

  if (cn.overrideLevels) {
    for (const level of cn.overrideLevels) {
      if (level.start !== undefined) {
        parts.push(
          `<w:lvlOverride w:ilvl="${level.num}"><w:startOverride w:val="${level.start}"/></w:lvlOverride>`,
        );
      } else {
        parts.push(`<w:lvlOverride w:ilvl="${level.num}"/>`);
      }
    }
  }

  parts.push("</w:num>");
  return parts.join("");
}

function stringifyLevel(opts: LevelsOptions): string {
  const children: string[] = [];

  children.push(`<w:start w:val="${decimalNumber(opts.start ?? 1)}"/>`);
  if (opts.format) children.push(`<w:numFmt w:val="${opts.format}"/>`);
  if (opts.lvlRestart !== undefined)
    children.push(`<w:lvlRestart w:val="${decimalNumber(opts.lvlRestart)}"/>`);
  if (opts.suffix) children.push(`<w:suff w:val="${opts.suffix}"/>`);
  if (opts.isLegalNumberingStyle) children.push("<w:isLgl/>");
  if (opts.text !== undefined || opts.textNull) {
    const lvlTextAttrs: string[] = [];
    if (opts.text !== undefined) lvlTextAttrs.push(`w:val="${opts.text}"`);
    if (opts.textNull) lvlTextAttrs.push('w:null="1"');
    children.push(`<w:lvlText ${lvlTextAttrs.join(" ")}/>`);
  }
  if (opts.lvlPicBulletId !== undefined)
    children.push(`<w:lvlPicBulletId w:val="${decimalNumber(opts.lvlPicBulletId)}"/>`);
  if (opts.legacy !== undefined) {
    const legacyAttrs: string[] = [`w:legacy="${(opts.legacy.enabled ?? true) ? 1 : 0}"`];
    if (opts.legacy.space !== undefined) legacyAttrs.push(`w:legacySpace="${opts.legacy.space}"`);
    if (opts.legacy.indent !== undefined)
      legacyAttrs.push(`w:legacyIndent="${opts.legacy.indent}"`);
    children.push(`<w:legacy ${legacyAttrs.join(" ")}/>`);
  }
  children.push(`<w:lvlJc w:val="${opts.alignment ?? AlignmentType.START}"/>`);

  // Paragraph/run properties — use compile-path pure string builders
  const pPrXml = stringifyParagraphProperties(opts.style && opts.style.paragraph).xml;
  const rPrXml = stringifyRunProperties(opts.style && opts.style.run);
  if (pPrXml) children.push(pPrXml);
  if (rPrXml) children.push(rPrXml);

  const lvlAttrs: string[] = [`w:ilvl="${decimalNumber(Math.min(opts.level, 9))}"`];
  // w15:tentative is optional; only emit when carried (matches sources that omit it).
  if (opts.w15Tentative !== undefined)
    lvlAttrs.push(`w15:tentative="${opts.w15Tentative ? 1 : 0}"`);
  if (opts.templateCode !== undefined) lvlAttrs.push(`w:tplc="${opts.templateCode}"`);
  if (opts.tentative !== undefined) lvlAttrs.push(`w:tentative="${opts.tentative ? 1 : 0}"`);

  return `<w:lvl ${lvlAttrs.join(" ")}>${children.join("")}</w:lvl>`;
}

// ── Parse (Element → NumberingOptions) ──

/**
 * Parse w:numbering element into NumberingOptions.
 */
export function parseNumberingDefinitions(
  el: Element,
  parseParagraphProperties: (el: Element, ctx: DocxReadContext) => Record<string, unknown>,
  ctx: DocxReadContext,
): NumberingOptions | undefined {
  const abstractNums = new Map<string, Element>();
  for (const child of el.elements ?? []) {
    if (child.name !== "w:abstractNum") continue;
    const id = attr(child, "w:abstractNumId");
    if (id !== undefined) abstractNums.set(id, child);
  }

  const numToAbstract = new Map<string, string>();
  for (const child of el.elements ?? []) {
    if (child.name !== "w:num") continue;
    const numId = attr(child, "w:numId");
    const abstractRef = findChild(child, "w:abstractNumId");
    const abstractId = abstractRef ? attr(abstractRef, "w:val") : undefined;
    if (numId !== undefined && abstractId !== undefined) {
      numToAbstract.set(numId, abstractId);
    }
  }

  const configs: NumberingOptions["config"] = [];

  for (const [numId, abstractId] of numToAbstract) {
    const abstractEl = abstractNums.get(abstractId);
    if (!abstractEl) continue;

    const levels: LevelsOptions[] = [];
    for (const child of abstractEl.elements ?? []) {
      if (child.name !== "w:lvl") continue;
      const levelOpts = parseLevelEl(child, parseParagraphProperties, ctx);
      if (levelOpts) levels.push(levelOpts);
    }

    if (levels.length > 0) {
      const extraOptions: AbstractNumberingExtraOptions = {};
      const nsidEl = findChild(abstractEl, "w:nsid");
      if (nsidEl) {
        const v = attr(nsidEl, "w:val");
        if (v) extraOptions.nsid = v;
      }
      const multiLevelTypeEl = findChild(abstractEl, "w:multiLevelType");
      if (multiLevelTypeEl) {
        const v = attr(multiLevelTypeEl, "w:val");
        if (v) extraOptions.multiLevelType = v;
      }
      const restartVal = attrBool(abstractEl, "w15:restartNumberingAfterBreak");
      if (restartVal !== undefined) extraOptions.restartNumberingAfterBreak = restartVal;
      const tmplEl = findChild(abstractEl, "w:tmpl");
      if (tmplEl) {
        const v = attr(tmplEl, "w:val");
        if (v) extraOptions.tmpl = v;
      }
      configs.push({ reference: `list_${numId}`, levels, extraOptions });
    }
  }

  if (configs.length === 0) return undefined;
  return { config: configs };
}

function parseLevelEl(
  el: Element,
  parseParagraphProperties: (el: Element, ctx: DocxReadContext) => Record<string, unknown>,
  ctx: DocxReadContext,
): LevelsOptions | undefined {
  const opts: Partial<LevelsOptions> = {};

  const level = attrNum(el, "w:ilvl");
  if (level !== undefined) opts.level = level;

  const start = findChild(el, "w:start");
  if (start) {
    const val = attrNum(start, "w:val");
    if (val !== undefined) opts.start = val;
  }

  const lvlRestart = findChild(el, "w:lvlRestart");
  if (lvlRestart) {
    const val = attrNum(lvlRestart, "w:val");
    if (val !== undefined) opts.lvlRestart = val;
  }

  const numFmt = findChild(el, "w:numFmt");
  if (numFmt) {
    const val = attr(numFmt, "w:val");
    if (val) opts.format = val as LevelsOptions["format"];
  }

  const suff = findChild(el, "w:suff");
  if (suff) {
    const val = attr(suff, "w:val");
    if (val) opts.suffix = val as LevelsOptions["suffix"];
  }

  if (findChild(el, "w:isLgl")) opts.isLegalNumberingStyle = true;

  const lvlText = findChild(el, "w:lvlText");
  if (lvlText) {
    const val = attr(lvlText, "w:val");
    if (val) opts.text = val;
    const isNull = attrBool(lvlText, "w:null");
    if (isNull) opts.textNull = isNull;
  }

  const lvlPicBulletId = findChild(el, "w:lvlPicBulletId");
  if (lvlPicBulletId) {
    const val = attrNum(lvlPicBulletId, "w:val");
    if (val !== undefined) opts.lvlPicBulletId = val;
  }

  // Legacy spacing (w:legacy/@w:legacy [required], @w:legacySpace, @w:legacyIndent)
  const legacyEl = findChild(el, "w:legacy");
  if (legacyEl) {
    const legacy: NonNullable<LevelsOptions["legacy"]> = {};
    const enabled = attrBool(legacyEl, "w:legacy");
    if (enabled !== undefined) legacy.enabled = enabled;
    const space = attrNum(legacyEl, "w:legacySpace");
    if (space !== undefined) legacy.space = space;
    const indent = attrNum(legacyEl, "w:legacyIndent");
    if (indent !== undefined) legacy.indent = indent;
    opts.legacy = legacy;
  }

  const lvlJc = findChild(el, "w:lvlJc");
  if (lvlJc) {
    const val = attr(lvlJc, "w:val");
    if (val) opts.alignment = val as LevelsOptions["alignment"];
  }

  // Level attributes (w:tplc templateCode; w:tentative; w15:tentative)
  const tplc = attr(el, "w:tplc");
  if (tplc) opts.templateCode = tplc;
  const tentative = attrBool(el, "w:tentative");
  if (tentative !== undefined) opts.tentative = tentative;
  const w15Tentative = attrBool(el, "w15:tentative");
  if (w15Tentative !== undefined) opts.w15Tentative = w15Tentative;

  // Run + paragraph properties — reuse the complete parse helpers for full
  // fidelity (stringifyLevel delegates to stringifyRunProperties /
  // stringifyParagraphProperties, so parse must use the matching readers).
  const style: NonNullable<LevelsOptions["style"]> = {};
  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    const runOpts = parseRunProperties(rPr);
    if (Object.keys(runOpts).length > 0) style.run = runOpts;
  }
  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    const paraOpts = parseParagraphProperties(pPr, ctx);
    if (Object.keys(paraOpts).length > 0) {
      style.paragraph = paraOpts;
    }
  }
  if (Object.keys(style).length > 0) opts.style = style;

  return Object.keys(opts).length > 0 ? (opts as LevelsOptions) : undefined;
}
