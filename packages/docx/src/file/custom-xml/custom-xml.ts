/**
 * Custom XML elements for WordprocessingML documents.
 *
 * Provides inline (CT_CustomXmlRun), block (CT_CustomXmlBlock),
 * row (CT_CustomXmlRow) and cell (CT_CustomXmlCell) custom XML
 * elements that wrap arbitrary content with XML element names and optional
 * properties.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_CustomXmlRun, CT_CustomXmlBlock,
 *   CT_CustomXmlRow, CT_CustomXmlCell
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import type { StructuredDocumentTagCell } from "@file/sdt";
import { TableCell } from "@file/table/table-cell/table-cell";
import type { TableRow } from "@file/table/table-row";
import { BuilderElement, XmlComponent } from "@file/xml-components";
import type { BaseXmlComponent } from "@file/xml-components";

/** Marker symbol for FileChild interface */
const FILE_CHILD = Symbol("fileChild");

// ── Options ──

/** Data binding configuration (CT_DataBinding) */
export interface CustomXmlDataBindingOptions {
  /** XPath expression for the data binding */
  xpath: string;
  /** Store item ID for the data binding */
  storeItemID: string;
  /** Namespace prefix mappings */
  prefixMappings?: string;
}

/** Custom attribute (CT_Attr) */
export interface CustomXmlAttributeOptions {
  name: string;
  val: string;
  uri?: string;
}

/** Custom XML properties (CT_CustomXmlPr) */
export interface CustomXmlPrOptions {
  /** Placeholder text */
  placeholder?: string;
  /** Custom attributes */
  attributes?: CustomXmlAttributeOptions[];
}

/** Options for inline custom XML (CT_CustomXmlRun) */
export interface CustomXmlRunOptions {
  /** XML element name (required) */
  element: string;
  /** Namespace URI */
  uri?: string;
  /** Properties (placeholder, data binding, attributes) */
  customXmlPr?: CustomXmlPrOptions;
  /** Inline content (runs, hyperlinks, etc.) */
  children?: BaseXmlComponent[];
}

/** Options for block-level custom XML (CT_CustomXmlBlock) */
export interface CustomXmlBlockOptions {
  /** XML element name (required) */
  element: string;
  /** Namespace URI */
  uri?: string;
  /** Properties (placeholder, data binding, attributes) */
  customXmlPr?: CustomXmlPrOptions;
  /** Block content (paragraphs, tables, etc.) */
  children?: FileChild[];
}

/** Options for row-level custom XML (CT_CustomXmlRow) */
export interface CustomXmlRowOptions {
  /** XML element name (required) */
  element: string;
  /** Namespace URI */
  uri?: string;
  /** Properties (placeholder, attributes) */
  customXmlPr?: CustomXmlPrOptions;
  /** Row content (TableRow children) */
  children?: TableRow[];
}

/** Options for cell-level custom XML (CT_CustomXmlCell) */
export interface CustomXmlCellOptions {
  /** XML element name (required) */
  element: string;
  /** Namespace URI */
  uri?: string;
  /** Properties (placeholder, attributes) */
  customXmlPr?: CustomXmlPrOptions;
  /** Cell content (TableCell or SdtCell children) */
  children?: (TableCell | StructuredDocumentTagCell)[];
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
  public fileChild = FILE_CHILD;

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

/**
 * Row-level custom XML element (CT_CustomXmlRow).
 *
 * Wraps table row content in a custom XML element.
 * Used inside `Table.rows` to annotate rows with custom XML metadata.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CustomXmlRow">
 *   <xsd:sequence>
 *     <xsd:element name="customXmlPr" type="CT_CustomXmlPr" minOccurs="0"/>
 *     <xsd:group ref="EG_ContentRowContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uri" type="ST_String"/>
 *   <xsd:attribute name="element" type="ST_XmlName" use="required"/>
 * </xsd:complexType>
 * ```
 */
export class CustomXmlRow extends XmlComponent {
  private rows: TableRow[];

  public constructor(options: CustomXmlRowOptions) {
    super("w:customXml");
    this.rows = options.children ?? [];

    const attrs: Record<string, string> = { "w:element": options.element };
    if (options.uri !== undefined) {
      attrs["w:uri"] = options.uri;
    }
    this.root.push({ _attr: attrs });

    if (options.customXmlPr !== undefined) {
      this.root.push(new CustomXmlPrComponent(options.customXmlPr));
    }

    for (const row of this.rows) {
      this.root.push(row);
    }
  }

  public get cellCount(): number {
    return Math.max(...this.rows.map((r) => r.cellCount), 0);
  }

  public get cells(): TableCell[] {
    return this.rows.flatMap((r) => r.cells);
  }

  /** @internal Used by Table to insert CONTINUE cells for vertical merge. */
  public addCellToColumnIndex(cell: TableCell, columnIndex: number): void {
    if (this.rows.length > 0) {
      this.rows[0].addCellToColumnIndex(cell, columnIndex);
    }
  }
}

/**
 * Cell-level custom XML element (CT_CustomXmlCell).
 *
 * Wraps table cell content in a custom XML element.
 * Used inside `TableRow.cells` to annotate cells with custom XML metadata.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CustomXmlCell">
 *   <xsd:sequence>
 *     <xsd:element name="customXmlPr" type="CT_CustomXmlPr" minOccurs="0"/>
 *     <xsd:group ref="EG_ContentCellContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uri" type="ST_String"/>
 *   <xsd:attribute name="element" type="ST_XmlName" use="required"/>
 * </xsd:complexType>
 * ```
 */
export class CustomXmlCell extends XmlComponent {
  public constructor(options: CustomXmlCellOptions) {
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
