/**
 * XML Mapping types for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd
 * CT_MapInfo, CT_Schema, CT_Map, CT_DataBinding,
 * CT_SingleXmlCells, CT_SingleXmlCell, CT_XmlCellPr, CT_XmlColumnPr, CT_XmlPr
 *
 * @module
 */

export interface SchemaOptions {
  /** Unique schema ID (required) */
  id: string;
  /** Schema reference */
  schemaRef?: string;
  /** Namespace URI */
  namespace?: string;
  /** Schema language */
  schemaLanguage?: string;
  /** Schema ID */
  schemaID?: string;
  /** Element form default */
  elementFormDefault?: string;
  /** Attribute form default */
  attributeFormDefault?: string;
}

export interface DataBindingOptions {
  /** Data binding name */
  dataBindingName?: string;
  /** Whether this is a file binding */
  fileBinding?: boolean;
  /** File binding name */
  fileBindingName?: string;
  /** Connection ID */
  connectionID?: number;
  /** Data binding load mode */
  dataBindingLoadMode?: number;
}

export interface MapOptions {
  /** Unique map ID (required) */
  id: number;
  /** Map name (required) */
  name: string;
  /** Root element name (required) */
  rootElement: string;
  /** Associated schema ID (required) */
  schemaID: string;
  /** Show import/export validation errors */
  showImportExportValidationErrors?: boolean;
  /** Append data */
  append?: boolean;
  /** Data binding load mode (on Map level) */
  dataBindingLoadMode?: number;
  /** Auto-fit columns */
  autoFit?: boolean;
  /** Whether this is a file binding (on Map level) */
  fileBinding?: boolean;
  /** File binding name (on Map level) */
  fileBindingName?: string;
  /** Preserve formatting */
  preserveFormat?: boolean;
  /** Preserve sort/autofilter layout */
  preserveSortAFLayout?: boolean;
  /** Data binding (child element) */
  dataBinding?: DataBindingOptions;
}

export interface MapInfoOptions {
  /** Selection namespaces (required) */
  selectionNamespaces: string;
  /** Schema definitions */
  schemas: SchemaOptions[];
  /** XML maps */
  maps: MapOptions[];
}

export interface XmlPrOptions {
  /** XML element name */
  xmlElement?: string;
  /** Associated map ID (required) */
  mapId: number;
  /** XPath expression (required) */
  xpath: string;
  /** XML data type (required) */
  xmlDataType: string;
}

export interface XmlCellPrOptions {
  /** Cell ID (required) */
  id: number;
  /** Unique name */
  uniqueName?: string;
  /** XML properties */
  xmlPr: XmlPrOptions;
}

export interface SingleXmlCellOptions {
  /** XML cell ID (required) */
  id: number;
  /** Cell reference, e.g. "A1" (required) */
  r: string;
  /** Connection ID (required) */
  connectionId: number;
  /** XML cell properties */
  xmlCellPr: XmlCellPrOptions;
}

export interface XmlColumnPrOptions {
  /** XPath expression (required) */
  xpath: string;
  /** XML data type (required) */
  xmlDataType: string;
  /** Associated map ID (required) */
  mapId: number;
  /** Denormalized (CT_XmlColumnPr @denormalized) */
  denormalized?: boolean;
}
