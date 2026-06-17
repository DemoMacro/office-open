/**
 * App (Extended) Properties module for WordprocessingML documents.
 *
 * Provides the AppPropertiesOptions interface and the appPropertiesDesc
 * descriptor for docProps/app.xml (CT_Properties).
 *
 * Reference: ISO-IEC29500-2_2016 shared-documentPropertiesExtended.xsd
 *
 * @module
 */

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml } from "@office-open/xml";

/** xsd:boolean lexical form — spec canonical form is "true"/"false" (Word's convention). */
const xsdBoolean = (value: boolean): string => (value ? "true" : "false");

/**
 * Options for docProps/app.xml extended properties (CT_Properties).
 *
 * Only the scalar (string/number/boolean) child elements are modelled.
 * Structured vector/blob elements (HeadingPairs, TitlesOfParts, HLinks,
 * DigSig, PresentationFormat) are intentionally omitted.
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
}

/** Subset of AppPropertiesOptions accepted by stringify. */
export type AppPropertiesInput = AppPropertiesOptions;

export const appPropertiesDesc: CustomDescriptor<AppPropertiesInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    // xsd:all is unordered, but emit in schema order for readability.
    const p: string[] = [
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    ];

    // Schema order: Template, Manager, Company, Pages, Words, Characters,
    // PresentationFormat, Lines, Paragraphs, Slides, Notes, TotalTime,
    // HiddenSlides, MMClips, ScaleCrop, HeadingPairs, TitlesOfParts,
    // LinksUpToDate, CharactersWithSpaces, SharedDoc, HyperlinkBase, HLinks,
    // HyperlinksChanged, DigSig, Application, AppVersion, DocSecurity.
    if (opts.template !== undefined) p.push(`<Template>${escapeXml(opts.template)}</Template>`);
    if (opts.manager !== undefined) p.push(`<Manager>${escapeXml(opts.manager)}</Manager>`);
    if (opts.company !== undefined) p.push(`<Company>${escapeXml(opts.company)}</Company>`);
    if (opts.pages !== undefined) p.push(`<Pages>${opts.pages}</Pages>`);
    if (opts.words !== undefined) p.push(`<Words>${opts.words}</Words>`);
    if (opts.characters !== undefined) p.push(`<Characters>${opts.characters}</Characters>`);
    if (opts.lines !== undefined) p.push(`<Lines>${opts.lines}</Lines>`);
    if (opts.paragraphs !== undefined) p.push(`<Paragraphs>${opts.paragraphs}</Paragraphs>`);
    if (opts.slides !== undefined) p.push(`<Slides>${opts.slides}</Slides>`);
    if (opts.notes !== undefined) p.push(`<Notes>${opts.notes}</Notes>`);
    if (opts.totalTime !== undefined) p.push(`<TotalTime>${opts.totalTime}</TotalTime>`);
    if (opts.hiddenSlides !== undefined)
      p.push(`<HiddenSlides>${opts.hiddenSlides}</HiddenSlides>`);
    if (opts.mmClips !== undefined) p.push(`<MMClips>${opts.mmClips}</MMClips>`);
    if (opts.scaleCrop !== undefined)
      p.push(`<ScaleCrop>${xsdBoolean(opts.scaleCrop)}</ScaleCrop>`);
    if (opts.linksUpToDate !== undefined)
      p.push(`<LinksUpToDate>${xsdBoolean(opts.linksUpToDate)}</LinksUpToDate>`);
    if (opts.charactersWithSpaces !== undefined)
      p.push(`<CharactersWithSpaces>${opts.charactersWithSpaces}</CharactersWithSpaces>`);
    if (opts.sharedDoc !== undefined)
      p.push(`<SharedDoc>${xsdBoolean(opts.sharedDoc)}</SharedDoc>`);
    if (opts.hyperlinkBase !== undefined)
      p.push(`<HyperlinkBase>${escapeXml(opts.hyperlinkBase)}</HyperlinkBase>`);
    if (opts.hyperlinksChanged !== undefined)
      p.push(`<HyperlinksChanged>${xsdBoolean(opts.hyperlinksChanged)}</HyperlinksChanged>`);
    if (opts.application !== undefined)
      p.push(`<Application>${escapeXml(opts.application)}</Application>`);
    if (opts.appVersion !== undefined)
      p.push(`<AppVersion>${escapeXml(opts.appVersion)}</AppVersion>`);
    if (opts.docSecurity !== undefined) p.push(`<DocSecurity>${opts.docSecurity}</DocSecurity>`);

    p.push("</Properties>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: AppPropertiesOptions = {};
    for (const child of el.elements ?? []) {
      if (typeof child.name !== "string") continue;
      const text = child.elements?.[0]?.text;
      switch (child.name) {
        case "Template":
          if (typeof text === "string") result.template = text;
          break;
        case "Manager":
          if (typeof text === "string") result.manager = text;
          break;
        case "Company":
          if (typeof text === "string") result.company = text;
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
        case "ScaleCrop":
          if (typeof text === "string") result.scaleCrop = text === "1" || text === "true";
          break;
        case "LinksUpToDate":
          if (typeof text === "string") result.linksUpToDate = text === "1" || text === "true";
          break;
        case "CharactersWithSpaces":
          if (typeof text === "string") result.charactersWithSpaces = Number(text);
          break;
        case "SharedDoc":
          if (typeof text === "string") result.sharedDoc = text === "1" || text === "true";
          break;
        case "HyperlinkBase":
          if (typeof text === "string") result.hyperlinkBase = text;
          break;
        case "HyperlinksChanged":
          if (typeof text === "string") result.hyperlinksChanged = text === "1" || text === "true";
          break;
        case "Application":
          if (typeof text === "string") result.application = text;
          break;
        case "AppVersion":
          if (typeof text === "string") result.appVersion = text;
          break;
        case "DocSecurity":
          if (typeof text === "string") result.docSecurity = Number(text);
          break;
      }
    }
    return result;
  },
};
