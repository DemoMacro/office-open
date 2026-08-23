/**
 * Extended (App) Properties module — shared OPC part (docProps/app.xml).
 *
 * Format-agnostic: CT_Properties (ISO-IEC29500-2_2016 shared-documentPropertiesExtended.xsd)
 * is identical across docx/pptx/xlsx. Each package surfaces only the fields it populates
 * (Pages/Words for docx, Slides for pptx, etc.) — all fields are optional.
 *
 * @module
 */

// ── Descriptor ──

import { escapeXml } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import { parseOnOff } from "../util/values";
import { parseVector, stringifyStringVector, stringifyVariantVector } from "./variant-types";
import type { VariantValue } from "./variant-types";

/** xsd:boolean lexical form — spec canonical form is "true"/"false" (Word's convention). */
const xsdBoolean = (value: boolean): string => (value ? "true" : "false");

/**
 * Options for docProps/app.xml extended properties (CT_Properties).
 *
 * Property order follows the CT_Properties xsd:all sequence so the
 * emitted XML matches the reference schema ordering.
 */
export interface AppPropertiesOptions {
  /** Template name */
  template?: string;
  /** Manager name */
  manager?: string;
  /** Company name */
  company?: string;
  /** Page count */
  pages?: number;
  /** Word count */
  words?: number;
  /** Character count */
  characters?: number;
  /** Intended print/display format, e.g. "Custom" or "A4 Paper" */
  presentationFormat?: string;
  /** Line count */
  lines?: number;
  /** Paragraph count */
  paragraphs?: number;
  /** Notes count */
  notes?: number;
  /** Slides count */
  slides?: number;
  /** Total editing time (minutes) */
  totalTime?: number;
  /** Hidden slides count */
  hiddenSlides?: number;
  /** Multimedia clips count */
  mmClips?: number;
  /** Characters including spaces */
  charactersWithSpaces?: number;
  /** Document security level */
  docSecurity?: number;
  /** Hyperlink base URL */
  hyperlinkBase?: string;
  /**
   * Hyperlink metadata vector (HLinks) — Word's internal per-link records
   * (six entries per link). Round-trip only: fields are undocumented, so the
   * flat variant values are carried as-is.
   */
  hlinks?: VariantValue[];
  /** Application name */
  application?: string;
  /** Application version */
  appVersion?: string;
  /** Whether the document is scaled/cropped */
  scaleCrop?: boolean;
  /** Whether links are up to date */
  linksUpToDate?: boolean;
  /** Whether the document is shared */
  sharedDoc?: boolean;
  /** Whether hyperlinks changed */
  hyperlinksChanged?: boolean;
  /** Named part-group counters (Office always writes these) */
  headingPairs?: HeadingPairOptions[];
  /** Part titles (slide titles, sheet names, heading ranges) */
  titlesOfParts?: string[];
  /**
   * Emit the extended-properties vocabulary under the ap: prefix — the ISO
   * strict binding, where the source declares xmlns:ap explicitly instead of
   * using the namespace as the default. Round-trip only.
   */
  apPrefix?: true;
}

/** One HeadingPairs entry: a group name plus the number of parts it covers. */
export interface HeadingPairOptions {
  /** Group name (e.g. "Fonts Used", "Theme", "Slide Titles") */
  name: string;
  /** Number of parts in the group */
  count: number;
}

/** Subset of AppPropertiesOptions accepted by stringify. */
export type AppPropertiesInput = AppPropertiesOptions;

export const appPropertiesDesc: CustomDescriptor<AppPropertiesInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    // xsd:all is unordered, but emit in schema order for readability. The
    // ap-prefixed form (ISO strict round-trip) binds xmlns:ap explicitly;
    // vt: children keep their own prefix either way.
    const t = (name: string): string => (opts.apPrefix ? `ap:${name}` : name);
    const p: string[] = opts.apPrefix
      ? [
          '<ap:Properties xmlns:ap="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
        ]
      : [
          '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
        ];

    // Schema order: Template, Manager, Company, Pages, Words, Characters,
    // PresentationFormat, Lines, Paragraphs, Slides, Notes, TotalTime,
    // HiddenSlides, MMClips, ScaleCrop, HeadingPairs, TitlesOfParts,
    // LinksUpToDate, CharactersWithSpaces, SharedDoc, HyperlinkBase, HLinks,
    // HyperlinksChanged, DigSig, Application, AppVersion, DocSecurity.
    if (opts.template !== undefined)
      p.push(`<${t("Template")}>${escapeXml(opts.template)}</${t("Template")}>`);
    if (opts.manager !== undefined)
      p.push(`<${t("Manager")}>${escapeXml(opts.manager)}</${t("Manager")}>`);
    if (opts.company !== undefined)
      p.push(`<${t("Company")}>${escapeXml(opts.company)}</${t("Company")}>`);
    if (opts.pages !== undefined) p.push(`<${t("Pages")}>${opts.pages}</${t("Pages")}>`);
    if (opts.words !== undefined) p.push(`<${t("Words")}>${opts.words}</${t("Words")}>`);
    if (opts.characters !== undefined)
      p.push(`<${t("Characters")}>${opts.characters}</${t("Characters")}>`);
    if (opts.presentationFormat !== undefined)
      p.push(
        `<${t("PresentationFormat")}>${escapeXml(opts.presentationFormat)}</${t("PresentationFormat")}>`,
      );
    if (opts.lines !== undefined) p.push(`<${t("Lines")}>${opts.lines}</${t("Lines")}>`);
    if (opts.paragraphs !== undefined)
      p.push(`<${t("Paragraphs")}>${opts.paragraphs}</${t("Paragraphs")}>`);
    if (opts.slides !== undefined) p.push(`<${t("Slides")}>${opts.slides}</${t("Slides")}>`);
    if (opts.notes !== undefined) p.push(`<${t("Notes")}>${opts.notes}</${t("Notes")}>`);
    if (opts.totalTime !== undefined)
      p.push(`<${t("TotalTime")}>${opts.totalTime}</${t("TotalTime")}>`);
    if (opts.hiddenSlides !== undefined)
      p.push(`<${t("HiddenSlides")}>${opts.hiddenSlides}</${t("HiddenSlides")}>`);
    if (opts.mmClips !== undefined) p.push(`<${t("MMClips")}>${opts.mmClips}</${t("MMClips")}>`);
    if (opts.scaleCrop !== undefined)
      p.push(`<${t("ScaleCrop")}>${xsdBoolean(opts.scaleCrop)}</${t("ScaleCrop")}>`);
    if (opts.headingPairs !== undefined && opts.headingPairs.length > 0) {
      // vt:lpstr is required here: Excel refuses to open a file whose
      // HeadingPairs variants carry lpwstr.
      const items = opts.headingPairs
        .map(
          (pair) =>
            `<vt:variant><vt:lpstr>${escapeXml(pair.name)}</vt:lpstr></vt:variant>` +
            `<vt:variant><vt:i4>${pair.count}</vt:i4></vt:variant>`,
        )
        .join("");
      p.push(
        `<${t("HeadingPairs")}><vt:vector size="${opts.headingPairs.length * 2}" baseType="variant">` +
          `${items}</vt:vector></${t("HeadingPairs")}>`,
      );
    }
    if (opts.titlesOfParts !== undefined && opts.titlesOfParts.length > 0) {
      p.push(
        `<${t("TitlesOfParts")}>${stringifyStringVector(opts.titlesOfParts)}</${t("TitlesOfParts")}>`,
      );
    }
    if (opts.linksUpToDate !== undefined)
      p.push(`<${t("LinksUpToDate")}>${xsdBoolean(opts.linksUpToDate)}</${t("LinksUpToDate")}>`);
    if (opts.charactersWithSpaces !== undefined)
      p.push(
        `<${t("CharactersWithSpaces")}>${opts.charactersWithSpaces}</${t("CharactersWithSpaces")}>`,
      );
    if (opts.sharedDoc !== undefined)
      p.push(`<${t("SharedDoc")}>${xsdBoolean(opts.sharedDoc)}</${t("SharedDoc")}>`);
    if (opts.hyperlinkBase !== undefined)
      p.push(`<${t("HyperlinkBase")}>${escapeXml(opts.hyperlinkBase)}</${t("HyperlinkBase")}>`);
    if (opts.hlinks !== undefined && opts.hlinks.length > 0)
      p.push(`<${t("HLinks")}>${stringifyVariantVector(opts.hlinks)}</${t("HLinks")}>`);
    if (opts.hyperlinksChanged !== undefined)
      p.push(
        `<${t("HyperlinksChanged")}>${xsdBoolean(opts.hyperlinksChanged)}</${t("HyperlinksChanged")}>`,
      );
    if (opts.application !== undefined)
      p.push(`<${t("Application")}>${escapeXml(opts.application)}</${t("Application")}>`);
    if (opts.appVersion !== undefined)
      p.push(`<${t("AppVersion")}>${escapeXml(opts.appVersion)}</${t("AppVersion")}>`);
    if (opts.docSecurity !== undefined)
      p.push(`<${t("DocSecurity")}>${opts.docSecurity}</${t("DocSecurity")}>`);

    p.push(opts.apPrefix ? "</ap:Properties>" : "</Properties>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: AppPropertiesOptions = {};
    // ISO strict binds the vocabulary under an explicit ap: prefix — strip
    // any prefix so ap:Template matches the local-name cases, and remember
    // the binding so stringify re-emits the same form.
    if (el.name === "ap:Properties") result.apPrefix = true;
    for (const child of el.elements ?? []) {
      if (typeof child.name !== "string") continue;
      const text = child.elements?.[0]?.text;
      // String fields are presence-based: Word writes empty elements
      // ("<Company>\n</Company>", which the XML parser reduces to no text)
      // — capture "" so the element round-trips instead of being dropped.
      const str = String(text ?? "");
      const name = child.name.includes(":")
        ? child.name.slice(child.name.indexOf(":") + 1)
        : child.name;
      switch (name) {
        case "Template":
          result.template = str;
          break;
        case "Manager":
          result.manager = str;
          break;
        case "Company":
          result.company = str;
          break;
        case "Pages":
          if (typeof text === "string") result.pages = Number(text);
          break;
        case "Words":
          if (typeof text === "string") result.words = Number(text);
          break;
        case "Characters":
          if (typeof text === "string") result.characters = Number(text);
          break;
        case "PresentationFormat":
          result.presentationFormat = str;
          break;
        case "Lines":
          if (typeof text === "string") result.lines = Number(text);
          break;
        case "Paragraphs":
          if (typeof text === "string") result.paragraphs = Number(text);
          break;
        case "Slides":
          if (typeof text === "string") result.slides = Number(text);
          break;
        case "Notes":
          if (typeof text === "string") result.notes = Number(text);
          break;
        case "TotalTime":
          if (typeof text === "string") result.totalTime = Number(text);
          break;
        case "HiddenSlides":
          if (typeof text === "string") result.hiddenSlides = Number(text);
          break;
        case "MMClips":
          if (typeof text === "string") result.mmClips = Number(text);
          break;
        case "HeadingPairs": {
          const vector = child.elements?.find((e) => e.name === "vt:vector");
          const flat = parseVector(vector);
          const pairs: HeadingPairOptions[] = [];
          for (let i = 0; i + 1 < flat.length; i += 2) {
            const name = flat[i];
            const count = flat[i + 1];
            if (typeof name === "string" && typeof count === "number") {
              pairs.push({ name, count });
            }
          }
          if (pairs.length > 0) result.headingPairs = pairs;
          break;
        }
        case "TitlesOfParts": {
          const vector = child.elements?.find((e) => e.name === "vt:vector");
          const titles = parseVector(vector).filter((v): v is string => typeof v === "string");
          if (titles.length > 0) result.titlesOfParts = titles;
          break;
        }
        case "ScaleCrop":
          if (typeof text === "string") result.scaleCrop = parseOnOff(text) ?? false;
          break;
        case "LinksUpToDate":
          if (typeof text === "string") result.linksUpToDate = parseOnOff(text) ?? false;
          break;
        case "CharactersWithSpaces":
          if (typeof text === "string") result.charactersWithSpaces = Number(text);
          break;
        case "SharedDoc":
          if (typeof text === "string") result.sharedDoc = parseOnOff(text) ?? false;
          break;
        case "HyperlinkBase":
          result.hyperlinkBase = str;
          break;
        case "HLinks": {
          const vector = child.elements?.find((e) => e.name === "vt:vector");
          const values = parseVector(vector);
          if (values.length > 0) result.hlinks = values;
          break;
        }
        case "HyperlinksChanged":
          if (typeof text === "string") result.hyperlinksChanged = parseOnOff(text) ?? false;
          break;
        case "Application":
          result.application = str;
          break;
        case "AppVersion":
          result.appVersion = str;
          break;
        case "DocSecurity":
          if (typeof text === "string") result.docSecurity = Number(text);
          break;
      }
    }
    return result;
  },
};
