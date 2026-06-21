/**
 * Styles module for WordprocessingML documents.
 *
 * Pure collector — no XmlComponent inheritance.
 * Collects child elements as raw XML strings and serializes.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * @module
 */
import { attr, attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ParagraphStylePropertiesOptions } from "@parts/paragraph/properties";
import type { RunStylePropertiesOptions } from "@parts/paragraph/run/properties";
import { parseRunProperties } from "@parts/paragraph/run/run-parse";
import type {
  CharacterStyleOptions,
  ParagraphStyleOptions,
  DefaultStylesOptions,
  DocumentDefaultsOptions,
} from "@parts/styles/factory";
import {
  STYLE_ID_TO_DEFAULT_FIELD,
  stringifyCharacterStyle,
  stringifyParagraphStyle,
} from "@parts/styles/factory";

import type { DocxReadContext } from "../../context";
import { stringifyElement } from "../../util/stringify-element";

/**
 * Options for configuring document styles.
 */
export interface StylesOptions {
  /** Default styles for document, headings, and common elements */
  default?: DefaultStylesOptions;
  /** Initial namespace attributes for the styles root element */
  initialAttributes?: Record<string, string>;
  /** Array of raw XML style elements */
  importedStyles?: { _raw: string }[];
  /** Array of custom paragraph style definitions */
  paragraphStyles?: (ParagraphStyleOptions & { id: string })[];
  /** Array of custom character style definitions */
  characterStyles?: (CharacterStyleOptions & { id: string })[];
  /**
   * Verbatim `<w:latentStyles>` XML (CT_LatentStyles + lsdException list).
   * When set (from parse), generate uses it in place of the default factory's
   * latent styles so the full latent-style table round-trips.
   */
  latentStylesXml?: string;
  /**
   * Verbatim `<w:docDefaults>` XML. When set (from parse), generate uses it in
   * place of the default factory's document defaults so rPrDefault/pPrDefault
   * round-trip byte-for-byte.
   */
  docDefaultsXml?: string;
}

/**
 * Extract the `w:styleId` from a raw `<w:style>` XML string, if present.
 *
 * Returns `undefined` for non-style elements (docDefaults / latentStyles),
 * whose raw XML carries no `w:styleId` attribute.
 */
export function extractStyleId(raw: string): string | undefined {
  const m = raw.match(/w:styleId="([^"]+)"/);
  return m ? m[1] : undefined;
}

/**
 * Represents the styles definitions in a WordprocessingML document.
 *
 * Pure collector — no XmlComponent inheritance.
 * Collects child elements as raw XML strings and serializes.
 */
export class Styles {
  private attributes: Record<string, string> = {};
  private parts: string[] = [];

  public constructor(options: StylesOptions) {
    if (options.initialAttributes) {
      this.attributes = options.initialAttributes;
    }

    // styleIds explicitly redefined via paragraphStyles/characterStyles take
    // precedence over importedStyles (user definitions override builtins) —
    // skip those imported entries to avoid duplicate styleId in the output.
    const customStyleIds = new Set<string>();
    for (const s of options.paragraphStyles ?? []) customStyleIds.add(s.id);
    for (const s of options.characterStyles ?? []) customStyleIds.add(s.id);

    if (options.importedStyles) {
      for (const style of options.importedStyles) {
        if (!style._raw) continue;
        if (customStyleIds.size > 0) {
          const id = extractStyleId(style._raw);
          if (id && customStyleIds.has(id)) continue;
        }
        this.parts.push(style._raw);
      }
    }

    if (options.paragraphStyles) {
      for (const style of options.paragraphStyles) {
        this.parts.push(
          stringifyParagraphStyle({
            id: style.id,
            name: style.name ?? style.id,
            aliases: style.aliases,
            basedOn: style.basedOn,
            next: style.next,
            link: style.link,
            autoRedefine: style.autoRedefine,
            quickFormat: style.quickFormat,
            semiHidden: style.semiHidden,
            uiPriority: style.uiPriority,
            unhideWhenUsed: style.unhideWhenUsed,
            locked: style.locked,
            personal: style.personal,
            personalCompose: style.personalCompose,
            personalReply: style.personalReply,
            paragraph: style.paragraph,
            run: style.run,
          }),
        );
      }
    }

    if (options.characterStyles) {
      for (const style of options.characterStyles) {
        this.parts.push(
          stringifyCharacterStyle({
            id: style.id,
            name: style.name ?? style.id,
            aliases: style.aliases,
            basedOn: style.basedOn,
            link: style.link,
            autoRedefine: style.autoRedefine,
            semiHidden: style.semiHidden,
            uiPriority: style.uiPriority,
            unhideWhenUsed: style.unhideWhenUsed,
            locked: style.locked,
            personal: style.personal,
            personalCompose: style.personalCompose,
            personalReply: style.personalReply,
            run: style.run,
          }),
        );
      }
    }
  }

  /**
   * Serialize to word/styles.xml content (with XML declaration).
   */
  public serialize(): string {
    const attrParts: string[] = [];
    for (const [k, v] of Object.entries(this.attributes)) {
      if (v !== undefined && v !== null) attrParts.push(` ${k}="${v}"`);
    }

    const attrs = attrParts.join("");
    const body = this.parts.join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles${attrs}>${body}</w:styles>`;
  }
}

// ── Parse helpers (for descriptor pipeline) ──

/** Style IDs generated by DefaultStylesFactory — skip these during parsing. */
const BUILTIN_STYLE_IDS = new Set([
  "Title",
  "Heading1",
  "Heading2",
  "Heading3",
  "Heading4",
  "Heading5",
  "Heading6",
  "Strong",
  "ListParagraph",
  "Hyperlink",
  "FootnoteText",
  "FootnoteTextChar",
  "EndnoteText",
  "EndnoteTextChar",
  "FootnoteReference",
  "EndnoteReference",
]);

/**
 * Build a cache of style elements keyed by styleId.
 */
export function buildStyleCache(stylesEl: Element | undefined): Map<string, Element> {
  const cache = new Map<string, Element>();
  if (!stylesEl) return cache;

  for (const child of stylesEl.elements ?? []) {
    if (child.name !== "w:style") continue;
    const styleId = attr(child, "w:styleId");
    if (styleId) {
      cache.set(styleId, child);
    }
  }

  return cache;
}

/**
 * Build a cache of abstract numbering elements keyed by abstractNumId.
 */
export function buildNumberingCache(numberingEl: Element | undefined): Map<string, Element> {
  const cache = new Map<string, Element>();
  if (!numberingEl) return cache;

  for (const child of numberingEl.elements ?? []) {
    if (child.name !== "w:abstractNum") continue;
    const abstractNumId = attr(child, "w:abstractNumId");
    if (abstractNumId !== undefined) {
      cache.set(abstractNumId, child);
    }
  }

  return cache;
}

/** Parsed style — concrete shape (not Record) with a transient _type marker. */
interface ParsedStyle {
  _type?: string;
  id?: string;
  default?: boolean;
  customStyle?: string;
  name?: string;
  aliases?: string;
  basedOn?: string;
  next?: string;
  link?: string;
  uiPriority?: number;
  quickFormat?: boolean;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  autoRedefine?: boolean;
  locked?: boolean;
  personal?: boolean;
  personalCompose?: boolean;
  personalReply?: boolean;
  paragraph?: ParagraphStylePropertiesOptions;
  run?: RunStylePropertiesOptions;
}

/**
 * Parse w:styles element into StylesOptions.
 *
 * Skips built-in styles that DefaultStylesFactory already generates,
 * keeping only user-defined custom styles for round-trip fidelity.
 */
export function parseStyleDefinitions(
  el: Element,
  parseParagraphProperties: (el: Element, ctx: DocxReadContext) => Record<string, unknown>,
  ctx: DocxReadContext,
): StylesOptions | undefined {
  const opts: StylesOptions = {};
  const paragraphStyles: (ParagraphStyleOptions & { id: string })[] = [];
  const characterStyles: (CharacterStyleOptions & { id: string })[] = [];

  for (const child of el.elements ?? []) {
    if (child.name === "w:docDefaults") {
      const defOpts = parseDocDefaults(child, parseParagraphProperties, ctx);
      if (defOpts) opts.default = defOpts;
      // Also capture verbatim for byte-exact round-trip (the structured path
      // may re-order or default-fill rPrDefault/pPrDefault children).
      opts.docDefaultsXml = stringifyElement(child);
    } else if (child.name === "w:latentStyles") {
      // Preserve the latent-style table verbatim so the full lsdException list
      // round-trips (the default factory only emits ~20 entries).
      opts.latentStylesXml = stringifyElement(child);
    } else if (child.name === "w:style") {
      const styleOpts = parseStyleElement(child, parseParagraphProperties, ctx);
      // Skip styles without a type or styleId — both are required to be useful.
      if (!styleOpts?._type || !styleOpts.id) continue;
      // Capture every style verbatim so the source's exact definitions
      // (numeric IDs, custom builtin formatting) round-trip instead of being
      // dropped + regenerated by the default factory.
      (opts.importedStyles ??= []).push({ _raw: stringifyElement(child) });
      // Map overridable builtins (Heading1-9, Title, Subtitle, ListParagraph,
      // Quote, Strong, Emphasis) into default.<field> so users can read and
      // edit them after parse; generate re-emits the edited version via the
      // factory (default.<field> wins over the verbatim entry).
      const defaultField = STYLE_ID_TO_DEFAULT_FIELD[styleOpts.id];
      if (defaultField) {
        const { _type: _omitType, id: _omitId, ...rest } = styleOpts;
        opts.default ??= {};
        (opts.default as Record<string, unknown>)[defaultField] = rest;
        continue;
      }
      // Skip structured parse of non-overridable builtins (footnote/endnote/
      // hyperlink/...); the factory regenerates them for scratch docs and the
      // raw capture above preserves them for round-trip.
      if (BUILTIN_STYLE_IDS.has(styleOpts.id)) continue;

      const type = styleOpts._type;
      delete styleOpts._type;

      if (type === "paragraph") {
        paragraphStyles.push(styleOpts as ParagraphStyleOptions & { id: string });
      } else if (type === "character") {
        characterStyles.push(styleOpts as CharacterStyleOptions & { id: string });
      }
    }
  }

  if (paragraphStyles.length > 0) opts.paragraphStyles = paragraphStyles;
  if (characterStyles.length > 0) opts.characterStyles = characterStyles;

  return Object.keys(opts).length > 0 ? opts : undefined;
}

function parseDocDefaults(
  el: Element,
  parseParagraphProperties: (el: Element, ctx: DocxReadContext) => Record<string, unknown>,
  ctx: DocxReadContext,
): DefaultStylesOptions | undefined {
  const document: DocumentDefaultsOptions = {};

  const rPrDefault = findChild(el, "w:rPrDefault");
  if (rPrDefault) {
    const rPr = findChild(rPrDefault, "w:rPr");
    if (rPr) {
      const runDefaults = parseRunProperties(rPr);
      if (Object.keys(runDefaults).length > 0) document.run = runDefaults;
    }
  }

  const pPrDefault = findChild(el, "w:pPrDefault");
  if (pPrDefault) {
    const pPr = findChild(pPrDefault, "w:pPr");
    if (pPr) {
      // Reuse the full paragraph-properties reader (stringifyDocDefaults uses
      // stringifyParagraphProperties) instead of only reading spacing, so jc/
      // ind/etc. round-trip too.
      const paraDefaults = parseParagraphProperties(pPr, ctx);
      if (Object.keys(paraDefaults).length > 0) {
        document.paragraph = paraDefaults as unknown as ParagraphStylePropertiesOptions;
      }
    }
  }

  return Object.keys(document).length > 0 ? { document } : undefined;
}

function parseStyleElement(
  el: Element,
  parseParagraphProperties: (el: Element, ctx: DocxReadContext) => Record<string, unknown>,
  ctx: DocxReadContext,
): ParsedStyle | undefined {
  const opts: ParsedStyle = {};

  const type = attr(el, "w:type");
  if (type) opts._type = type;

  const id = attr(el, "w:styleId");
  if (id) opts.id = id;

  if (attrBool(el, "w:default")) opts.default = true;
  if (attrBool(el, "w:customStyle")) opts.customStyle = "1";

  const nameEl = findChild(el, "w:name");
  if (nameEl) {
    const name = attr(nameEl, "w:val");
    if (name) opts.name = name;
  }

  const basedOn = findChild(el, "w:basedOn");
  if (basedOn) {
    const val = attr(basedOn, "w:val");
    if (val) opts.basedOn = val;
  }

  const next = findChild(el, "w:next");
  if (next) {
    const val = attr(next, "w:val");
    if (val) opts.next = val;
  }

  const link = findChild(el, "w:link");
  if (link) {
    const val = attr(link, "w:val");
    if (val) opts.link = val;
  }

  const uiPriority = findChild(el, "w:uiPriority");
  if (uiPriority) {
    const val = attrNum(uiPriority, "w:val");
    if (val !== undefined) opts.uiPriority = val;
  }

  if (findChild(el, "w:qFormat")) opts.quickFormat = true;
  if (findChild(el, "w:semiHidden")) opts.semiHidden = true;
  if (findChild(el, "w:unhideWhenUsed")) opts.unhideWhenUsed = true;
  if (findChild(el, "w:autoRedefine")) opts.autoRedefine = true;
  if (findChild(el, "w:locked")) opts.locked = true;
  if (findChild(el, "w:personal")) opts.personal = true;
  if (findChild(el, "w:personalCompose")) opts.personalCompose = true;
  if (findChild(el, "w:personalReply")) opts.personalReply = true;

  const aliases = findChild(el, "w:aliases");
  if (aliases) {
    const val = attr(aliases, "w:val");
    if (val) opts.aliases = val;
  }

  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    const paraOpts = parseParagraphProperties(pPr, ctx);
    if (Object.keys(paraOpts).length > 0) {
      opts.paragraph = paraOpts as unknown as ParagraphStylePropertiesOptions;
    }
  }

  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    const runOpts = parseRunProperties(rPr);
    if (Object.keys(runOpts).length > 0) opts.run = runOpts;
  }

  return opts;
}
