import type { BibliographyOptions } from "@parts/bibliography";
/**
 * Core Properties module for WordprocessingML documents.
 *
 * Provides the DocumentOptions interface for document metadata.
 * XML generation is handled by the descriptor pipeline (corePropertiesDesc).
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import type { ContentTypesInput } from "@parts/contenttypes";
import type { EmbeddedFontOptions } from "@parts/fonts/font-table";
import type { GlossaryDocumentOptions } from "@parts/glossary-document";
import type { CommentsOptions } from "@parts/paragraph/run/comment-run";
import type { HyphenationOptions } from "@parts/settings";
import type { CompatibilityOptions } from "@parts/settings/compatibility";
import type {
  DocumentProtectionOptions,
  SettingsOptions,
  WriteProtectionOptions,
} from "@parts/settings/settings";
import type { SectionOptions } from "@shared/section";

import type { CustomPropertyOptions } from "./custom-properties";
import type { DocumentBackgroundOptions } from "./document";
import type { NumberingOptions } from "./numbering";
import type { ParagraphOptions } from "./paragraph";
import type { StylesOptions } from "./styles";
import type { WebSettingsOptions } from "./web-settings";

/**
 * Document-level feature toggles parsed from settings.xml.
 *
 * @property trackRevisions - Track changes
 * @property updateFields - Update fields on open
 * @property documentProtection - Document write protection
 */
export interface FeaturesOptions {
  trackRevisions?: boolean;
  updateFields?: boolean;
  documentProtection?: DocumentProtectionOptions;
}

/**
 * Options for configuring document properties.
 *
 * @property sections - Document section configurations
 * @property title - Document title
 * @property subject - Document subject
 * @property creator - Document creator/author
 * @property keywords - Document keywords for searchability
 * @property description - Document description
 * @property lastModifiedBy - User who last modified the document
 * @property revision - Revision number
 * @property externalStyles - External stylesheet reference
 * @property styles - Document styles configuration
 * @property numbering - Numbering configuration
 * @property comments - Document comments configuration
 * @property bibliography - Document bibliography sources
 * @property footnotes - Document footnotes
 * @property background - Document background settings
 * @property features - Document features like track changes
 * @property compatabilityModeVersion - Compatibility mode version
 * @property compatibility - Compatibility settings
 * @property customProperties - Custom document properties
 * @property evenAndOddHeaderAndFooters - Enable different headers/footers for even/odd pages
 * @property defaultTabStop - Default tab stop width
 * @property fonts - Font configurations
 * @property hyphenation - Hyphenation settings
 */
export interface DocumentOptions {
  sections: SectionOptions[];
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: number;
  externalStyles?: string;
  styles?: StylesOptions;
  numbering?: NumberingOptions;
  comments?: CommentsOptions;
  bibliography?: BibliographyOptions;
  footnotes?: Readonly<
    Record<
      string,
      {
        children: (ParagraphOptions | string)[];
      }
    >
  >;
  endnotes?: Readonly<
    Record<
      string,
      {
        children: (ParagraphOptions | string)[];
      }
    >
  >;
  background?: DocumentBackgroundOptions;
  features?: FeaturesOptions;
  compatabilityModeVersion?: number;
  compatibility?: CompatibilityOptions;
  customProperties?: CustomPropertyOptions[];
  evenAndOddHeaderAndFooters?: boolean;
  defaultTabStop?: number;
  fonts?: EmbeddedFontOptions[];
  hyphenation?: HyphenationOptions;
  /** Controls whether punctuation is compressed at line ends */
  characterSpacingControl?: "compressPunctuation" | "doNotCompress";
  /** Default document view mode */
  view?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
  /** Default zoom level (percentage) and type */
  zoom?: {
    percent?: number;
    val?: "none" | "fullPage" | "bestFit" | "textFit";
  };
  /** Write protection recommendation (not enforcement) */
  writeProtection?: WriteProtectionOptions;
  /** Whether to display the background shape in print layout */
  displayBackgroundShape?: boolean;
  /** Whether to embed TrueType fonts in the document */
  embedTrueTypeFonts?: boolean;
  /** Whether to embed system fonts in the document */
  embedSystemFonts?: boolean;
  /** Whether to save only a subset of the embedded fonts */
  saveSubsetFonts?: boolean;
  /** Document variables (key-value pairs stored in the document) */
  docVars?: { name: string; val: string }[];
  /** Theme color scheme remapping */
  colorSchemeMapping?: SettingsOptions["colorSchemeMapping"];
  /** Mail merge configuration */
  mailMerge?: SettingsOptions["mailMerge"];
  /** Glossary document — building blocks (Quick Parts) */
  glossary?: GlossaryDocumentOptions;
  /** Additional document settings passed through to the settings.xml part */
  settings?: SettingsOptions;
  /** Web settings for browser rendering (word/webSettings.xml) */
  webSettings?: WebSettingsOptions;
  /** Content types from [Content_Types].xml (parse path only) */
  contentTypes?: ContentTypesInput;
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml } from "@office-open/xml";

/** Subset of DocumentOptions used for core properties XML. */
export interface CorePropertiesInput {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: number;
}

export const corePropertiesDesc: CustomDescriptor<CorePropertiesInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const now = new Date().toISOString();
    const p: string[] = [
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"' +
        ' xmlns:dcterms="http://purl.org/dc/terms/"' +
        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"' +
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    ];
    if (opts.title) p.push(`<dc:title>${escapeXml(opts.title)}</dc:title>`);
    if (opts.subject) p.push(`<dc:subject>${escapeXml(opts.subject)}</dc:subject>`);
    if (opts.creator) p.push(`<dc:creator>${escapeXml(opts.creator)}</dc:creator>`);
    if (opts.keywords) p.push(`<cp:keywords>${escapeXml(opts.keywords)}</cp:keywords>`);
    if (opts.description) p.push(`<dc:description>${escapeXml(opts.description)}</dc:description>`);
    if (opts.lastModifiedBy)
      p.push(`<cp:lastModifiedBy>${escapeXml(opts.lastModifiedBy)}</cp:lastModifiedBy>`);
    if (opts.revision !== undefined) p.push(`<cp:revision>${opts.revision}</cp:revision>`);
    p.push(`<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`);
    p.push(`<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`);
    p.push("</cp:coreProperties>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};
    for (const child of el.elements ?? []) {
      if (typeof child.name !== "string") continue;
      const text = child.elements?.[0]?.text;
      if (typeof text !== "string") continue;
      switch (child.name) {
        case "dc:title":
          result.title = text;
          break;
        case "dc:subject":
          result.subject = text;
          break;
        case "dc:creator":
          result.creator = text;
          break;
        case "cp:keywords":
          result.keywords = text;
          break;
        case "dc:description":
          result.description = text;
          break;
        case "cp:lastModifiedBy":
          result.lastModifiedBy = text;
          break;
        case "cp:revision":
          result.revision = Number(text);
          break;
      }
    }
    return result as Record<string, unknown>;
  },
};
