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
import type { SectionChild } from "@shared/section";

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
 */
export interface CustomXmlRunOptions {
  /** XML element name (required) */
  element: string;
  /** Namespace URI */
  uri?: string;
  /** Properties (placeholder, data binding, attributes) */
  customXmlPr?: CustomXmlPropertiesOptions;
}

/**
 * Options for block-level custom XML (CT_CustomXmlBlock).
 * Wraps block content (paragraphs, tables, …); lives in EG_BlockLevelElts.
 */
export type CustomXmlBlockOptions = CustomXmlRunOptions & {
  /** Block content (paragraphs, tables, etc.) */
  children?: SectionChild[];
};

/**
 * Options for row-level custom XML (CT_CustomXmlRow).
 * Wraps one or more table rows; lives in EG_ContentRowContent alongside w:tr.
 */
export type CustomXmlRowOptions = CustomXmlRunOptions & {
  /** Row content (TableRow children) */
  children?: TableRowOptions[];
};

/**
 * Options for cell-level custom XML (CT_CustomXmlCell).
 * Wraps one or more table cells; lives in EG_ContentCellContent alongside w:tc.
 */
export type CustomXmlCellOptions = CustomXmlRunOptions & {
  /** Cell content (TableCell children) */
  children?: TableCellOptions[];
};
