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
import type { ContentTypesInput, DataType } from "@office-open/core";
import type { BibliographyOptions } from "@parts/bibliography";
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

import type { AppPropertiesOptions } from "./app-properties";
import type { CustomPropertyOptions } from "./custom-properties";
import type { DocumentBackgroundOptions } from "./document";
import type { EndnoteSeparator } from "./endnotes/descriptor";
import type { FootnoteSeparator } from "./footnotes/descriptor";
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
 * @property lastPrinted - Last printed timestamp (W3CDTF), round-tripped from cp:lastPrinted
 * @property externalStyles - External stylesheet reference
 * @property styles - Document styles configuration
 * @property numbering - Numbering configuration
 * @property comments - Document comments configuration
 * @property bibliography - Document bibliography sources
 * @property footnotes - Document footnotes
 * @property background - Document background settings
 * @property features - Document features like track changes
 * @property compatibility - Compatibility settings
 * @property customProperties - Custom document properties
 * @property evenAndOddHeaderAndFooters - Enable different headers/footers for even/odd pages
 * @property defaultTabStop - Default tab stop width
 * @property fonts - Font configurations
 * @property hyphenation - Hyphenation settings
 */
export interface DocumentOptions extends CorePropertiesOptions {
  sections: SectionOptions[];
  externalStyles?: string;
  styles?: StylesOptions;
  numbering?: NumberingOptions;
  comments?: CommentsOptions;
  bibliography?: BibliographyOptions;
  footnotes?: Record<string, { children: (ParagraphOptions | string)[] }> & {
    /**
     * Separator footnote — id + content round-tripped verbatim from the source
     * so the generated id stays consistent with settings.footnotePr, which
     * references it. Omit for freshly generated documents (defaults apply).
     */
    separator?: FootnoteSeparator;
    /** Continuation separator footnote — round-tripped verbatim from the source. */
    continuationSeparator?: FootnoteSeparator;
  };
  endnotes?: Record<string, { children: (ParagraphOptions | string)[] }> & {
    /** Separator endnote — id + content round-tripped verbatim from the source. */
    separator?: EndnoteSeparator;
    /** Continuation separator endnote — round-tripped verbatim from the source. */
    continuationSeparator?: EndnoteSeparator;
  };
  background?: DocumentBackgroundOptions;
  features?: FeaturesOptions;
  compatibility?: CompatibilityOptions;
  customProperties?: CustomPropertyOptions[];
  evenAndOddHeaderAndFooters?: boolean;
  defaultTabStop?: number;
  fonts?: EmbeddedFontOptions[];
  hyphenation?: HyphenationOptions;
  /** Document conformance class (w:document/@w:conformance). */
  conformance?: "strict" | "transitional";
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
  /**
   * Parts carried verbatim from the source that generate() does not rebuild
   * (e.g. word/theme/*, customXml/*). Kept as raw bytes so their [Content_Types]
   * declarations stay valid and the package opens in Word. Media/fonts/headers
   * are rebuilt by the compiler and must NOT be listed here.
   */
  rawParts?: { path: string; data: DataType }[];
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
}

// ── Descriptor ──

import type { CorePropertiesOptions } from "@office-open/core";
import { buildCorePropertiesXmlString, parseCorePropsElement } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";

export const corePropertiesDesc: CustomDescriptor<CorePropertiesOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildCorePropertiesXmlString(opts);
  },

  parse(el, _ctx) {
    return parseCorePropsElement(el);
  },
};
