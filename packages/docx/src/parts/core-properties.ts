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
import type { ContentTypesInput, DataType, EncryptedContainerOptions } from "@office-open/core";
import type { BibliographyOptions } from "@parts/bibliography";
import type { CommentExtendedOptions } from "@parts/comments-extended";
import type { EmbeddedFontOptions } from "@parts/fonts/font-table";
import type { GlossaryDocumentOptions } from "@parts/glossary-document";
import type { CommentOptions } from "@parts/paragraph/run/comment-run";
import type { CommentPersonOptions } from "@parts/people";
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
 * @property styles - Document styles configuration
 * @property numbering - Numbering configuration
 * @property comments - Document comments (word/comments.xml)
 * @property people - Comment authors (word/people.xml)
 * @property commentsExtended - Comment metadata (word/commentsExtended.xml)
 * @property bibliography - Document bibliography sources
 * @property footnotes - Document footnotes
 * @property background - Document background settings
 * @property customProperties - Custom document properties
 * @property fonts - Font configurations
 * @property settings - Document settings (word/settings.xml)
 */
export interface DocumentOptions extends CorePropertiesOptions {
  sections: SectionOptions[];
  /**
   * The source file is an encrypted OOXML package (OLE2/CFB container).
   * Round-trip only: the plaintext needs the password, so the original bytes
   * are carried verbatim and generate() re-emits them unchanged — every
   * other field stays empty (`sections: []`). Mixing real content is
   * rejected — it would be silently dropped.
   */
  encrypted?: EncryptedContainerOptions;
  styles?: StylesOptions;
  numbering?: NumberingOptions;
  comments?: CommentOptions[];
  /**
   * Comment authors (word/people.xml). Word pairs each w15:person with
   * comments by exact author-string equality; omit for documents whose
   * comments carry no author registry.
   */
  people?: CommentPersonOptions[];
  /**
   * Comment metadata (word/commentsExtended.xml) — resolved state and reply
   * threading, keyed by the w14:paraId of each comment's first paragraph.
   */
  commentsExtended?: CommentExtendedOptions[];
  bibliography?: BibliographyOptions;
  /** User footnotes (word/footnotes.xml). `id` auto-assigns 1, 2, … when omitted. */
  footnotes?: FootnoteOptions[];
  /**
   * Separator footnotes — id + content round-tripped verbatim from the source
   * so the generated id stays consistent with settings.footnoteProperties,
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
  /** Document conformance class (w:document/`@w:conformance`). */
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
   * (e.g. word/theme/*, customXml/*, any unknown extension part). Collected
   * wholesale by the core passthrough pipeline — everything the model did not
   * absorb survives with bytes and content-type declaration intact. Parts the
   * compiler happens to rebuild under the same path win over the passthrough
   * copy (assembly order), so media/fonts/headers need no exclusion here.
   */
  rawParts?: { path: string; data: DataType; contentType?: string }[];
  /**
   * Relationships from rebuilt parts' source .rels that point at rawParts
   * (e.g. document.xml → theme/customXml). Re-emitted verbatim — target
   * unchanged (passthrough paths never move), fresh rId. Round-trip only.
   */
  passthroughRelationships?: {
    source: string;
    relationshipType: string;
    target: string;
    rId: string;
  }[];
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
