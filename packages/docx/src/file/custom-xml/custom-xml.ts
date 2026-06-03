/**
 * Custom XML elements for WordprocessingML documents.
 *
 * Provides inline (CT_CustomXmlRun) and block (CT_CustomXmlBlock) custom XML
 * elements that wrap arbitrary content with XML element names and optional
 * data bindings.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_CustomXmlRun, CT_CustomXmlBlock
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { BuilderElement, XmlComponent } from "@file/xml-components";
import type { BaseXmlComponent } from "@file/xml-components";

/** Marker symbol for FileChild interface */
const FILE_CHILD = Symbol("fileChild");

// ── Options ──

/** Data binding configuration (CT_DataBinding) */
export interface CustomXmlDataBindingOptions {
  /** XPath expression for the data binding */
  readonly xpath: string;
  /** Store item ID for the data binding */
  readonly storeItemID: string;
  /** Namespace prefix mappings */
  readonly prefixMappings?: string;
}

/** Custom attribute (CT_Attr) */
export interface CustomXmlAttributeOptions {
  readonly name: string;
  readonly val: string;
  readonly uri?: string;
}

/** Custom XML properties (CT_CustomXmlPr) */
export interface CustomXmlPrOptions {
  /** Placeholder text */
  readonly placeholder?: string;
  /** Custom attributes */
  readonly attributes?: readonly CustomXmlAttributeOptions[];
}

/** Options for inline custom XML (CT_CustomXmlRun) */
export interface CustomXmlRunOptions {
  /** XML element name (required) */
  readonly element: string;
  /** Namespace URI */
  readonly uri?: string;
  /** Properties (placeholder, data binding, attributes) */
  readonly customXmlPr?: CustomXmlPrOptions;
  /** Inline content (runs, hyperlinks, etc.) */
  readonly children?: readonly BaseXmlComponent[];
}

/** Options for block-level custom XML (CT_CustomXmlBlock) */
export interface CustomXmlBlockOptions {
  /** XML element name (required) */
  readonly element: string;
  /** Namespace URI */
  readonly uri?: string;
  /** Properties (placeholder, data binding, attributes) */
  readonly customXmlPr?: CustomXmlPrOptions;
  /** Block content (paragraphs, tables, etc.) */
  readonly children?: readonly FileChild[];
}

// ── Components ──

/**
 * Inline custom XML element (CT_CustomXmlRun).
 *
 * Wraps run-level content (TextRun, ImageRun, etc.) in a custom XML element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CustomXmlRun">
 *   <xsd:sequence>
 *     <xsd:element name="customXmlPr" type="CT_CustomXmlPr" minOccurs="0"/>
 *     <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uri" type="ST_String"/>
 *   <xsd:attribute name="element" type="ST_XmlName" use="required"/>
 * </xsd:complexType>
 * ```
 */
export class CustomXmlRun extends XmlComponent {
  public constructor(options: CustomXmlRunOptions) {
    super("w:customXml");

    const attrs: Record<string, string> = { "w:element": options.element };
    if (options.uri !== undefined) {
      attrs["w:uri"] = options.uri;
    }
    this.root.push({ _attr: attrs });

    if (options.customXmlPr !== undefined) {
      this.root.push(new CustomXmlPrComponent(options.customXmlPr));
    }

    if (options.children !== undefined) {
      for (const child of options.children) {
        this.root.push(child);
      }
    }
  }
}

/**
 * Block-level custom XML element (CT_CustomXmlBlock).
 *
 * Wraps block content (Paragraph, Table, etc.) in a custom XML element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CustomXmlBlock">
 *   <xsd:sequence>
 *     <xsd:element name="customXmlPr" type="CT_CustomXmlPr" minOccurs="0"/>
 *     <xsd:group ref="EG_ContentBlockContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uri" type="ST_String"/>
 *   <xsd:attribute name="element" type="ST_XmlName" use="required"/>
 * </xsd:complexType>
 * ```
 */
export class CustomXmlBlock extends XmlComponent {
  public readonly fileChild = FILE_CHILD;

  public constructor(options: CustomXmlBlockOptions) {
    super("w:customXml");

    const attrs: Record<string, string> = { "w:element": options.element };
    if (options.uri !== undefined) {
      attrs["w:uri"] = options.uri;
    }
    this.root.push({ _attr: attrs });

    if (options.customXmlPr !== undefined) {
      this.root.push(new CustomXmlPrComponent(options.customXmlPr));
    }

    if (options.children !== undefined) {
      for (const child of options.children) {
        this.root.push(child);
      }
    }
  }
}

// ── Internal ──

/** Custom XML properties (CT_CustomXmlPr) */
class CustomXmlPrComponent extends XmlComponent {
  public constructor(options: CustomXmlPrOptions) {
    super("w:customXmlPr");

    if (options.placeholder !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:placeholder",
          attributes: [{ key: "w:val", value: options.placeholder }],
        }),
      );
    }

    if (options.attributes !== undefined) {
      for (const attr of options.attributes) {
        const attrItems: { key: string; value: string }[] = [
          { key: "w:name", value: attr.name },
          { key: "w:val", value: attr.val },
        ];
        if (attr.uri !== undefined) {
          attrItems.push({ key: "w:uri", value: attr.uri });
        }
        this.root.push(new BuilderElement({ name: "w:attr", attributes: attrItems }));
      }
    }
  }
}
