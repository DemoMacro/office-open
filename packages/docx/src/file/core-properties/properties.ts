import type { BibliographyOptions } from "@file/bibliography";
/**
 * Core Properties module for WordprocessingML documents.
 *
 * Provides support for document metadata based on Dublin Core properties.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import type { EmbeddedFontOptions } from "@file/fonts/font-table";
import type { CommentsOptions } from "@file/paragraph/run/comment-run";
import type { HyphenationOptions } from "@file/settings";
import type { CompatibilityOptions } from "@file/settings/compatibility";
import type {
  DocumentProtectionOptions,
  SettingsOptions,
  WriteProtectionOptions,
} from "@file/settings/settings";
import { XmlComponent, stringContainerObj } from "@file/xml-components";
import { dateTimeValue } from "@util/values";

import type { CustomPropertyOptions } from "../custom-properties";
import type { DocumentBackgroundOptions } from "../document";
import { buildDocumentAttributes } from "../document/document-attributes";
import type { SectionOptions } from "../file";
import type { NumberingOptions } from "../numbering";
import type { ParagraphOptions, Paragraph } from "../paragraph";
import type { StylesOptions } from "../styles";

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
export interface PropertiesOptions {
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
        children: (Paragraph | ParagraphOptions | string)[];
      }
    >
  >;
  endnotes?: Readonly<
    Record<
      string,
      {
        children: (Paragraph | ParagraphOptions | string)[];
      }
    >
  >;
  background?: DocumentBackgroundOptions;
  features?: {
    trackRevisions?: boolean;
    updateFields?: boolean;
    documentProtection?: DocumentProtectionOptions;
  };
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
  glossary?: import("../glossary/glossary-document").GlossaryDocumentOptions;
  /** Additional document settings passed through to the settings.xml part */
  settings?: import("../settings/settings").SettingsOptions;
  /** Web settings for browser rendering (word/webSettings.xml) */
  webSettings?: import("../web-settings/web-settings").WebSettingsOptions;
}

/**
 * Represents the core properties of a WordprocessingML document.
 *
 * Core properties contain document metadata based on Dublin Core elements,
 * including title, subject, creator, keywords, description, and modification tracking.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * ## XSD Schema
 * ```xml
 * <xs:complexType name="CT_CoreProperties">
 *   <xs:all>
 *     <xs:element name="category" minOccurs="0" maxOccurs="1" type="xs:string"/>
 *     <xs:element name="contentStatus" minOccurs="0" maxOccurs="1" type="xs:string"/>
 *     <xs:element ref="dcterms:created" minOccurs="0" maxOccurs="1"/>
 *     <xs:element ref="dc:creator" minOccurs="0" maxOccurs="1"/>
 *     <xs:element ref="dc:description" minOccurs="0" maxOccurs="1"/>
 *     <xs:element ref="dc:identifier" minOccurs="0" maxOccurs="1"/>
 *     <xs:element name="keywords" minOccurs="0" maxOccurs="1" type="CT_Keywords"/>
 *     <xs:element ref="dc:language" minOccurs="0" maxOccurs="1"/>
 *     <xs:element name="lastModifiedBy" minOccurs="0" maxOccurs="1" type="xs:string"/>
 *     <xs:element name="lastPrinted" minOccurs="0" maxOccurs="1" type="xs:dateTime"/>
 *     <xs:element ref="dcterms:modified" minOccurs="0" maxOccurs="1"/>
 *     <xs:element name="revision" minOccurs="0" maxOccurs="1" type="xs:string"/>
 *     <xs:element ref="dc:subject" minOccurs="0" maxOccurs="1"/>
 *     <xs:element ref="dc:title" minOccurs="0" maxOccurs="1"/>
 *     <xs:element name="version" minOccurs="0" maxOccurs="1" type="xs:string"/>
 *   </xs:all>
 * </xs:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const coreProps = new CoreProperties({
 *   title: "My Document",
 *   subject: "Sample Document",
 *   creator: "John Doe",
 *   keywords: "docx, example",
 *   description: "A sample document",
 *   lastModifiedBy: "Jane Doe",
 *   revision: 1
 * });
 * ```
 */
export class CoreProperties extends XmlComponent {
  public constructor(options: Omit<PropertiesOptions, "sections">) {
    super("cp:coreProperties");
    this.root.push(buildDocumentAttributes(["cp", "dc", "dcterms", "dcmitype", "xsi"]));
    if (options.title) {
      this.root.push(stringContainerObj("dc:title", options.title));
    }
    if (options.subject) {
      this.root.push(stringContainerObj("dc:subject", options.subject));
    }
    if (options.creator) {
      this.root.push(stringContainerObj("dc:creator", options.creator));
    }
    if (options.keywords) {
      this.root.push(stringContainerObj("cp:keywords", options.keywords));
    }
    if (options.description) {
      this.root.push(stringContainerObj("dc:description", options.description));
    }
    if (options.lastModifiedBy) {
      this.root.push(stringContainerObj("cp:lastModifiedBy", options.lastModifiedBy));
    }
    if (options.revision) {
      this.root.push(stringContainerObj("cp:revision", String(options.revision)));
    }
    this.root.push(new TimestampElement("dcterms:created"));
    this.root.push(new TimestampElement("dcterms:modified"));
  }
}

/**
 * Represents a timestamp element (created or modified date).
 * Uses W3C DateTime Format (dcterms:W3CDTF) for dates.
 */
class TimestampElement extends XmlComponent {
  public constructor(name: string) {
    super(name);
    this.root.push({ _attr: { "xsi:type": "dcterms:W3CDTF" } });
    this.root.push(dateTimeValue(new Date()));
  }
}
