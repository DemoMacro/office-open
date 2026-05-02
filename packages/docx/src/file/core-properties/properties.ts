import type { IBibliographyOptions } from "@file/bibliography";
/**
 * Core Properties module for WordprocessingML documents.
 *
 * Provides support for document metadata based on Dublin Core properties.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import type { FontOptions } from "@file/fonts/font-table";
import type { ICommentsOptions } from "@file/paragraph/run/comment-run";
import type { IHyphenationOptions } from "@file/settings";
import type { ICompatibilityOptions } from "@file/settings/compatibility";
import { StringContainer, XmlAttributeComponent, XmlComponent } from "@file/xml-components";
import { dateTimeValue } from "@util/values";

import type { ICustomPropertyOptions } from "../custom-properties";
import type { IDocumentBackgroundOptions } from "../document";
import { DocumentAttributes } from "../document/document-attributes";
import type { ISectionOptions } from "../file";
import type { INumberingOptions } from "../numbering";
import type { Paragraph } from "../paragraph";
import type { IStylesOptions } from "../styles";

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
export interface IPropertiesOptions {
    readonly sections: readonly ISectionOptions[];
    readonly title?: string;
    readonly subject?: string;
    readonly creator?: string;
    readonly keywords?: string;
    readonly description?: string;
    readonly lastModifiedBy?: string;
    readonly revision?: number;
    readonly externalStyles?: string;
    readonly styles?: IStylesOptions;
    readonly numbering?: INumberingOptions;
    readonly comments?: ICommentsOptions;
    readonly bibliography?: IBibliographyOptions;
    readonly footnotes?: Readonly<
        Record<
            string,
            {
                readonly children: readonly Paragraph[];
            }
        >
    >;
    readonly endnotes?: Readonly<
        Record<
            string,
            {
                readonly children: readonly Paragraph[];
            }
        >
    >;
    readonly background?: IDocumentBackgroundOptions;
    readonly features?: {
        readonly trackRevisions?: boolean;
        readonly updateFields?: boolean;
        readonly documentProtection?: import("@file/settings/settings").IDocumentProtectionOptions;
    };
    readonly compatabilityModeVersion?: number;
    readonly compatibility?: ICompatibilityOptions;
    readonly customProperties?: readonly ICustomPropertyOptions[];
    readonly evenAndOddHeaderAndFooters?: boolean;
    readonly defaultTabStop?: number;
    readonly fonts?: readonly FontOptions[];
    readonly hyphenation?: IHyphenationOptions;
    /** Controls whether punctuation is compressed at line ends */
    readonly characterSpacingControl?: "compressPunctuation" | "doNotCompress";
    /** Default document view mode */
    readonly view?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
    /** Default zoom level (percentage) and type */
    readonly zoom?: {
        readonly percent?: number;
        readonly val?: "none" | "fullPage" | "bestFit" | "textFit";
    };
    /** Write protection recommendation (not enforcement) */
    readonly writeProtection?: import("@file/settings/settings").IWriteProtectionOptions;
    /** Whether to display the background shape in print layout */
    readonly displayBackgroundShape?: boolean;
    /** Whether to embed TrueType fonts in the document */
    readonly embedTrueTypeFonts?: boolean;
    /** Whether to embed system fonts in the document */
    readonly embedSystemFonts?: boolean;
    /** Whether to save only a subset of the embedded fonts */
    readonly saveSubsetFonts?: boolean;
    /** Document variables (key-value pairs stored in the document) */
    readonly docVars?: readonly { readonly name: string; readonly val: string }[];
    /** Theme color scheme remapping */
    readonly colorSchemeMapping?: import("@file/settings/settings").ISettingsOptions["colorSchemeMapping"];
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
    public constructor(options: Omit<IPropertiesOptions, "sections">) {
        super("cp:coreProperties");
        this.root.push(new DocumentAttributes(["cp", "dc", "dcterms", "dcmitype", "xsi"]));
        if (options.title) {
            this.root.push(new StringContainer("dc:title", options.title));
        }
        if (options.subject) {
            this.root.push(new StringContainer("dc:subject", options.subject));
        }
        if (options.creator) {
            this.root.push(new StringContainer("dc:creator", options.creator));
        }
        if (options.keywords) {
            this.root.push(new StringContainer("cp:keywords", options.keywords));
        }
        if (options.description) {
            this.root.push(new StringContainer("dc:description", options.description));
        }
        if (options.lastModifiedBy) {
            this.root.push(new StringContainer("cp:lastModifiedBy", options.lastModifiedBy));
        }
        if (options.revision) {
            this.root.push(new StringContainer("cp:revision", String(options.revision)));
        }
        this.root.push(new TimestampElement("dcterms:created"));
        this.root.push(new TimestampElement("dcterms:modified"));
    }
}

/**
 * Attributes for timestamp elements in core properties.
 * Specifies the W3C DateTime Format type for timestamps.
 */
class TimestampElementProperties extends XmlAttributeComponent<{ readonly type: string }> {
    protected readonly xmlKeys = { type: "xsi:type" };
}

/**
 * Represents a timestamp element (created or modified date).
 * Uses W3C DateTime Format (dcterms:W3CDTF) for dates.
 */
class TimestampElement extends XmlComponent {
    public constructor(name: string) {
        super(name);
        this.root.push(
            new TimestampElementProperties({
                type: "dcterms:W3CDTF",
            }),
        );
        this.root.push(dateTimeValue(new Date()));
    }
}
