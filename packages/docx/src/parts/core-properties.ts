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
import type { CommentOptions } from "@parts/paragraph/run/comment-run";
import type { SettingsOptions } from "@parts/settings/settings";
import type { SectionOptions } from "@shared/section";

import type { AppPropertiesOptions } from "./app-properties";
import type { CustomPropertyOptions } from "./custom-properties";
import type { DocumentBackgroundOptions } from "./document";
import type { EndnoteOptions, EndnoteSeparator } from "./endnotes/descriptor";
import type { FootnoteOptions, FootnoteSeparator } from "./footnotes/descriptor";
import type { NumberingOptions } from "./numbering";
import type { StylesOptions } from "./styles";
import type { WebSettingsOptions } from "./web-settings";

/**
 * Options for configuring document properties.
 *
 * All settings.xml content is configured through the single {@link settings}
 * entry (SettingsOptions) — mirroring the OOXML part structure, where
 * word/settings.xml is a standalone part and the document body carries none
 * of it.
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
 * @property comments - Document comments (word/comments.xml)
 * @property bibliography - Document bibliography sources
 * @property footnotes - Document footnotes
 * @property background - Document background settings
 * @property customProperties - Custom document properties
 * @property fonts - Font configurations
 * @property settings - Document settings (word/settings.xml)
 */
export interface DocumentOptions extends CorePropertiesOptions {
  sections: SectionOptions[];
  externalStyles?: string;
  styles?: StylesOptions;
  numbering?: NumberingOptions;
  comments?: CommentOptions[];
  bibliography?: BibliographyOptions;
  /** User footnotes (word/footnotes.xml). `id` auto-assigns 1, 2, … when omitted. */
  footnotes?: FootnoteOptions[];
  /**
   * Separator footnotes — id + content round-tripped verbatim from the source
   * so the generated id stays consistent with settings.footnotePr, which
   * references it. Omit for freshly generated documents (defaults apply).
   */
  footnoteSeparators?: {
    separator?: FootnoteSeparator;
    continuationSeparator?: FootnoteSeparator;
  };
  /** User endnotes (word/endnotes.xml). `id` auto-assigns 1, 2, … when omitted. */
  endnotes?: EndnoteOptions[];
  /** Separator endnotes — round-tripped verbatim from the source. */
  endnoteSeparators?: {
    separator?: EndnoteSeparator;
    continuationSeparator?: EndnoteSeparator;
  };
  background?: DocumentBackgroundOptions;
  customProperties?: CustomPropertyOptions[];
  fonts?: EmbeddedFontOptions[];
  /** Document conformance class (w:document/@w:conformance). */
  conformance?: "strict" | "transitional";
  /** Glossary document — building blocks (Quick Parts) */
  glossary?: GlossaryDocumentOptions;
  /** Document settings (word/settings.xml). */
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
