/**
 * Hyperlink module for WordprocessingML documents.
 *
 * This module provides hyperlink functionality for internal and external links.
 *
 * Reference: http://officeopenxml.com/WPhyperlink.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";

import type { ParagraphChild } from "../paragraph";

interface HyperlinkAttributesProperties {
  readonly id?: string;
  readonly anchor?: string;
  readonly history: number;
  readonly tooltip?: string;
  readonly tgtFrame?: string;
  readonly docLocation?: string;
}

/**
 * Hyperlink type enumeration.
 *
 * Defines the types of hyperlinks supported in WordprocessingML documents.
 *
 * @publicApi
 */
export const HyperlinkType = {
  /** Internal hyperlink to a bookmark within the document */
  INTERNAL: "INTERNAL",
  /** External hyperlink to a URL outside the document */
  EXTERNAL: "EXTERNAL",
} as const;

/**
 * Options for creating an internal hyperlink.
 *
 * @property children - Array of paragraph children (usually TextRun elements) that form the hyperlink text
 * @property anchor - Name of the bookmark to link to within the document
 */
export interface InternalHyperlinkOptions {
  /** Array of paragraph children that form the hyperlink text */
  readonly children: readonly ParagraphChild[];
  /** Name of the bookmark to link to within the document */
  readonly anchor: string;
  /** Screen tip text shown when hovering over the hyperlink */
  readonly tooltip?: string;
}

/**
 * Options for creating an external hyperlink.
 *
 * @property children - Array of paragraph children (usually TextRun elements) that form the hyperlink text
 * @property link - URL to link to outside the document
 * @property tooltip - Screen tip text shown when hovering over the hyperlink
 * @property tgtFrame - Target frame for the hyperlink (e.g., "_blank", "_self")
 */
export interface ExternalHyperlinkOptions {
  /** Array of paragraph children that form the hyperlink text */
  readonly children: readonly ParagraphChild[];
  /** URL to link to outside the document */
  readonly link: string;
  /** Screen tip text shown when hovering over the hyperlink */
  readonly tooltip?: string;
  /** Target frame for the hyperlink (e.g., "_blank", "_self") */
  readonly tgtFrame?: string;
}

/**
 * Represents a concrete hyperlink in a WordprocessingML document.
 *
 * This class is the low-level implementation of hyperlinks used internally.
 * Use InternalHyperlink or ExternalHyperlink for creating hyperlinks in documents.
 *
 * Reference: http://officeopenxml.com/WPhyperlink.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="hyperlink" type="CT_Hyperlink"/>
 *
 * <xsd:complexType name="CT_Hyperlink">
 *   <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
 *   <xsd:attribute name="tgtFrame" type="s:ST_String" use="optional"/>
 *   <xsd:attribute name="tooltip" type="s:ST_String" use="optional"/>
 *   <xsd:attribute name="docLocation" type="s:ST_String" use="optional"/>
 *   <xsd:attribute name="history" type="s:ST_OnOff" use="optional"/>
 *   <xsd:attribute name="anchor" type="s:ST_String" use="optional"/>
 *   <xsd:attribute ref="r:id"/>
 * </xsd:complexType>
 * ```
 */
export interface ConcreteHyperlinkOptions {
  readonly anchor?: string;
  readonly tooltip?: string;
  readonly tgtFrame?: string;
}

export class ConcreteHyperlink extends XmlComponent {
  public readonly linkId: string;

  public constructor(
    children: readonly ParagraphChild[],
    relationshipId: string,
    options?: ConcreteHyperlinkOptions,
  ) {
    super("w:hyperlink");

    this.linkId = relationshipId;

    const anchor = options?.anchor;
    const tooltip = options?.tooltip;
    const tgtFrame = options?.tgtFrame;

    const props: HyperlinkAttributesProperties = {
      anchor: anchor ? anchor : undefined,
      history: 1,
      id: !anchor ? `rId${this.linkId}` : undefined,
      tooltip,
      tgtFrame,
    };

    const attr: Record<string, string | number> = {};
    if (props.anchor !== undefined) {
      attr["w:anchor"] = props.anchor;
    }
    attr["w:history"] = props.history;
    if (props.id !== undefined) {
      attr["r:id"] = props.id;
    }
    if (props.tooltip !== undefined) {
      attr["w:tooltip"] = props.tooltip;
    }
    if (props.tgtFrame !== undefined) {
      attr["w:tgtFrame"] = props.tgtFrame;
    }
    this.root.push({ _attr: attr });
    for (const child of children) {
      this.root.push(child);
    }
  }
}

/**
 * Represents an internal hyperlink to a bookmark within the document.
 *
 * Internal hyperlinks use the anchor attribute to reference a bookmark by name.
 * The bookmark must exist in the document for the hyperlink to function.
 *
 * Reference: http://officeopenxml.com/WPhyperlink.php
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="hyperlink" type="CT_Hyperlink"/>
 *
 * <xsd:complexType name="CT_Hyperlink">
 *   <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
 *   <xsd:attribute name="anchor" type="s:ST_String" use="optional"/>
 *   <xsd:attribute name="history" type="s:ST_OnOff" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create a bookmark
 * new Bookmark({
 *   id: "section1",
 *   children: [new TextRun("Section 1")],
 * });
 *
 * // Link to the bookmark
 * new InternalHyperlink({
 *   children: [new TextRun({ text: "Go to Section 1", style: "Hyperlink" })],
 *   anchor: "section1",
 * });
 * ```
 */
export class InternalHyperlink extends ConcreteHyperlink {
  public constructor(options: InternalHyperlinkOptions) {
    super(options.children, uniqueId(), { anchor: options.anchor, tooltip: options.tooltip });
  }
}

/**
 * Represents an external hyperlink to a URL outside the document.
 *
 * External hyperlinks create a relationship to an external resource (URL).
 * The relationship is created during document preparation and the hyperlink
 * is converted to a ConcreteHyperlink with the relationship ID.
 *
 * Reference: http://officeopenxml.com/WPhyperlink.php
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="hyperlink" type="CT_Hyperlink"/>
 *
 * <xsd:complexType name="CT_Hyperlink">
 *   <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
 *   <xsd:attribute ref="r:id"/>
 *   <xsd:attribute name="history" type="s:ST_OnOff" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new ExternalHyperlink({
 *   children: [new TextRun({ text: "Visit Example", style: "Hyperlink" })],
 *   link: "https://example.com",
 * });
 * ```
 */
export class ExternalHyperlink extends XmlComponent {
  public constructor(public readonly options: ExternalHyperlinkOptions) {
    super("w:externalHyperlink");
  }
}
