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
import { documentNamespaceAttributes } from "@parts/document/document-attributes";
import { AlignmentType } from "@parts/paragraph";
import type { ParagraphPropertiesOptions } from "@parts/paragraph/properties";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";

import type { DocxReadContext } from "../../context";
import { stringifyElement } from "../../util/stringify-element";
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
    /** Verbatim w:pict element XML (CT_Picture) */
    pict?: string;
    /** Verbatim w:drawing element XML (CT_Drawing) — alternative to pict */
    drawing?: string;
  }[];
}

/** Namespace attributes for w:numbering (pre-computed constant). */
const NUMBERING_ATTRS =
  documentNamespaceAttributes([
    "wpc",
    "mc",
    "o",
    "r",
    "m",
    "v",
    "wp14",
    "wp",
    "w10",
    "w",
    "w14",
    "w15",
    "wpg",
    "wpi",
    "wne",
    "wps",
  ]) + ' mc:Ignorable="w14 w15 wp14"';

/** Default bullet levels (9 levels: 0-8). */
const DEFAULT_BULLET_LEVELS: LevelsOptions[] = [
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 0,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: convertInchesToTwip(0.5) } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 1,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: convertInchesToTwip(1) } },
    text: "○",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 2,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 2160 } },
    text: "■",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 3,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 2880 } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 4,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 3600 } },
    text: "○",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 5,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 4320 } },
    text: "■",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 6,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 5040 } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 7,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 5760 } },
    text: "●",
  },
  {
    alignment: AlignmentType.LEFT,
    format: LevelFormat.BULLET,
    level: 8,
    paragraph: { indent: { hanging: convertInchesToTwip(0.25), left: 6480 } },
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
  private _numPicBullets?: { numPicBulletId: number; pict?: string; drawing?: string }[];

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
        const inner = bullet.pict ?? bullet.drawing;
        if (inner) {
          parts.push(
            `<w:numPicBullet w:numPicBulletId="${bullet.numPicBulletId}">${inner}</w:numPicBullet>`,
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

/** Extra options for abstract numbering (w:abstractNum attributes + w15 restart). */
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

// String-valued w:abstractNum child elements: XML tag → options key.
const ABSTRACT_EXTRA_PROPS = [
  ["w:nsid", "nsid"],
  ["w:multiLevelType", "multiLevelType"],
  ["w:tmpl", "tmpl"],
  ["w:name", "name"],
  ["w:styleLink", "styleLink"],
  ["w:numStyleLink", "numStyleLink"],
] as const;

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
  if (opts.levelRestart !== undefined)
    children.push(`<w:lvlRestart w:val="${decimalNumber(opts.levelRestart)}"/>`);
  if (opts.paragraphStyle !== undefined)
    children.push(`<w:pStyle w:val="${opts.paragraphStyle}"/>`);
  // CT_Lvl sequence: pStyle → isLgl → suff (XSD). isLgl must precede suff.
  if (opts.isLegalNumberingStyle) children.push("<w:isLgl/>");
  if (opts.suffix) children.push(`<w:suff w:val="${opts.suffix}"/>`);
  if (opts.text !== undefined || opts.textNull) {
    const lvlTextAttrs: string[] = [];
    if (opts.text !== undefined) lvlTextAttrs.push(`w:val="${opts.text}"`);
    if (opts.textNull) lvlTextAttrs.push('w:null="1"');
    children.push(`<w:lvlText ${lvlTextAttrs.join(" ")}/>`);
  }
  if (opts.levelPictureBulletId !== undefined)
    children.push(`<w:lvlPicBulletId w:val="${decimalNumber(opts.levelPictureBulletId)}"/>`);
  if (opts.legacy !== undefined) {
    const legacyAttrs: string[] = [`w:legacy="${(opts.legacy.enabled ?? true) ? 1 : 0}"`];
    if (opts.legacy.space !== undefined) legacyAttrs.push(`w:legacySpace="${opts.legacy.space}"`);
    if (opts.legacy.indent !== undefined)
      legacyAttrs.push(`w:legacyIndent="${opts.legacy.indent}"`);
    children.push(`<w:legacy ${legacyAttrs.join(" ")}/>`);
  }
  children.push(`<w:lvlJc w:val="${opts.alignment ?? AlignmentType.START}"/>`);

  // Paragraph/run properties — use compile-path pure string builders
  const pPrXml = stringifyParagraphProperties(opts.paragraph).xml;
  const rPrXml = stringifyRunProperties(opts.run);
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
  parseParagraphProperties: (
    el: Element,
    ctx: DocxReadContext,
  ) => Partial<ParagraphPropertiesOptions>,
  ctx: DocxReadContext,
): NumberingOptions | undefined {
  // Picture bullets (w:numPicBullet — choice of w:pict | w:drawing)
  const numPicBullets: NonNullable<NumberingOptions["numPicBullets"]> = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "w:numPicBullet") continue;
    const id = attrNum(child, "w:numPicBulletId");
    const bullet: { numPicBulletId: number; pict?: string; drawing?: string } = {
      numPicBulletId: id ?? 0,
    };
    const pictEl = findChild(child, "w:pict");
    if (pictEl) bullet.pict = stringifyElement(pictEl);
    const drawingEl = findChild(child, "w:drawing");
    if (drawingEl) bullet.drawing = stringifyElement(drawingEl);
    numPicBullets.push(bullet);
  }

  // numIdMacAtCleanup (w:numIdMacAtCleanup)
  const cleanupEl = findChild(el, "w:numIdMacAtCleanup");
  const numIdMacAtCleanup = cleanupEl ? attrNum(cleanupEl, "w:val") : undefined;

  const abstractNums = new Map<string, Element>();
  for (const child of el.elements ?? []) {
    if (child.name !== "w:abstractNum") continue;
    const id = attr(child, "w:abstractNumId");
    if (id !== undefined) abstractNums.set(id, child);
  }

  // Concrete num instances: numId → abstractId + the num element. The element
  // is kept so its lvlOverride/startOverride can be read — a concrete num may
  // re-pin a level's start, and dropping the override silently reverts the
  // list's restart numbering on round-trip.
  const numEntries: { numId: string; abstractId: string; numEl: Element }[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "w:num") continue;
    const numId = attr(child, "w:numId");
    const abstractRef = findChild(child, "w:abstractNumId");
    const abstractId = abstractRef ? attr(abstractRef, "w:val") : undefined;
    if (numId !== undefined && abstractId !== undefined) {
      numEntries.push({ numId, abstractId, numEl: child });
    }
  }

  const configs: NumberingOptions["config"] = [];

  for (const { numId, abstractId, numEl } of numEntries) {
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
      for (const [tag, key] of ABSTRACT_EXTRA_PROPS) {
        const childEl = findChild(abstractEl, tag);
        const v = childEl ? attr(childEl, "w:val") : undefined;
        if (v) extraOptions[key] = v;
      }
      const restartVal = attrBool(abstractEl, "w15:restartNumberingAfterBreak");
      if (restartVal !== undefined) extraOptions.restartNumberingAfterBreak = restartVal;
      // Apply per-instance level start overrides. The abstract level defines
      // the default start; a concrete num may re-pin it via lvlOverride, and
      // dropping it silently reverts the list's restart numbering.
      for (const overrideEl of numEl.elements ?? []) {
        if (overrideEl.name !== "w:lvlOverride") continue;
        const ilvl = attrNum(overrideEl, "w:ilvl");
        if (ilvl === undefined) continue;
        const startOverrideEl = findChild(overrideEl, "w:startOverride");
        if (!startOverrideEl) continue;
        const val = attrNum(startOverrideEl, "w:val");
        if (val === undefined) continue;
        const level = levels.find((l) => l.level === ilvl);
        if (level) level.start = val;
      }
      configs.push({ reference: `list_${numId}`, levels, extraOptions });
    }
  }

  if (configs.length === 0 && numPicBullets.length === 0 && numIdMacAtCleanup === undefined) {
    return undefined;
  }
  const result: NumberingOptions = { config: configs };
  if (numPicBullets.length > 0) result.numPicBullets = numPicBullets;
  if (numIdMacAtCleanup !== undefined) result.numIdMacAtCleanup = numIdMacAtCleanup;
  return result;
}

function parseLevelEl(
  el: Element,
  parseParagraphProperties: (
    el: Element,
    ctx: DocxReadContext,
  ) => Partial<ParagraphPropertiesOptions>,
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
    if (val !== undefined) opts.levelRestart = val;
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

  const pStyle = findChild(el, "w:pStyle");
  if (pStyle) {
    const val = attr(pStyle, "w:val");
    if (val) opts.paragraphStyle = val;
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
    if (val !== undefined) opts.levelPictureBulletId = val;
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
  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    const runOpts = parseRunProperties(rPr);
    if (Object.keys(runOpts).length > 0) opts.run = runOpts;
  }
  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    const paraOpts = parseParagraphProperties(pPr, ctx);
    if (Object.keys(paraOpts).length > 0) {
      opts.paragraph = paraOpts;
    }
  }

  return Object.keys(opts).length > 0 ? (opts as LevelsOptions) : undefined;
}
