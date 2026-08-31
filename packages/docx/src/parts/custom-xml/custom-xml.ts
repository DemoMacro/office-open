import type { TableCellOptions } from "@parts/table/table-cell/table-cell";
import type { TableRowOptions } from "@parts/table/table-row";
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
import type { BlockContentChild } from "@shared/section";

// ── Options ──

/** Custom attribute (CT_Attr) */
export interface CustomXmlAttributeOptions {
  name: string;
  val: string;
  uri?: string;
}

/** Custom XML properties (CT_CustomXmlProperties) */
export interface CustomXmlPropertiesOptions {
  /** Placeholder text */
  placeholder?: string;
  /** Custom attributes */
  attributes?: CustomXmlAttributeOptions[];
}

/**
 * Base shape shared by all four custom XML levels
 * (CT_CustomXmlRun / CT_CustomXmlBlock / CT_CustomXmlRow / CT_CustomXmlCell):
 * the w:customXml element name + optional namespace URI + optional properties.
 * Each level extends this with its own `children` content type.
 *
 * Word deletes inline `w:customXml` markup on open — round-trip only; use
 * content controls (`w:sdt`) or customXml parts for new content.
 */
export interface CustomXmlRunOptions {
  /** XML element name (required) */
  element: string;
  /** Namespace URI */
  uri?: string;
  /** Properties (placeholder, data binding, attributes) */
  properties?: CustomXmlPropertiesOptions;
}

/**
 * Block-level custom XML (CT_CustomXmlBlock), wrapping paragraphs/tables.
 * Word deletes inline w:customXml on open — round-trip only; use `w:sdt`
 * for new content.
 */
export type CustomXmlBlockOptions = CustomXmlRunOptions & {
  /** Block content (paragraphs, tables, etc.) — no altChunk (XSD EG_ContentBlockContent). */
  children?: BlockContentChild[];
};

/**
 * Row-level custom XML (CT_CustomXmlRow), wrapping table rows.
 * Word deletes inline w:customXml on open — round-trip only; use `w:sdt`
 * for new content.
 */
export type CustomXmlRowOptions = CustomXmlRunOptions & {
  /** Row content (TableRow children) */
  children?: TableRowOptions[];
};

/**
 * Cell-level custom XML (CT_CustomXmlCell), wrapping table cells.
 * Word deletes inline w:customXml on open — round-trip only; use `w:sdt`
 * for new content.
 */
export type CustomXmlCellOptions = CustomXmlRunOptions & {
  /** Cell content (TableCell children) */
  children?: TableCellOptions[];
};
