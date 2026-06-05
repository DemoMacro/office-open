/**
 * XML Mapping elements — produces XML spreadsheet mapping elements.
 *
 * Reference: OOXML transitional, sml.xsd
 * CT_MapInfo, CT_Schema, CT_Map, CT_DataBinding,
 * CT_SingleXmlCells, CT_SingleXmlCell, CT_XmlCellPr, CT_XmlColumnPr, CT_XmlPr
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs } from "@office-open/xml";

// ── Options ──

export interface SchemaOptions {
  /** Unique schema ID (required) */
  readonly id: string;
  /** Schema reference */
  readonly schemaRef?: string;
  /** Namespace URI */
  readonly namespace?: string;
  /** Schema language */
  readonly schemaLanguage?: string;
  /** Schema ID */
  readonly schemaID?: string;
  /** Element form default */
  readonly elementFormDefault?: string;
  /** Attribute form default */
  readonly attributeFormDefault?: string;
}

export interface DataBindingOptions {
  /** Data binding name */
  readonly dataBindingName?: string;
  /** Whether this is a file binding */
  readonly fileBinding?: boolean;
  /** File binding name */
  readonly fileBindingName?: string;
  /** Connection ID */
  readonly connectionID?: number;
  /** Data binding load mode */
  readonly dataBindingLoadMode?: number;
}

export interface MapOptions {
  /** Unique map ID (required) */
  readonly id: number;
  /** Map name (required) */
  readonly name: string;
  /** Root element name (required) */
  readonly rootElement: string;
  /** Associated schema ID (required) */
  readonly schemaID: string;
  /** Show import/export validation errors */
  readonly showImportExportValidationErrors?: boolean;
  /** Append data */
  readonly append?: boolean;
  /** Data binding load mode (on Map level) */
  readonly dataBindingLoadMode?: number;
  /** Auto-fit columns */
  readonly autoFit?: boolean;
  /** Whether this is a file binding (on Map level) */
  readonly fileBinding?: boolean;
  /** File binding name (on Map level) */
  readonly fileBindingName?: string;
  /** Preserve formatting */
  readonly preserveFormat?: boolean;
  /** Preserve sort/autofilter layout */
  readonly preserveSortAFLayout?: boolean;
  /** Data binding (child element) */
  readonly dataBinding?: DataBindingOptions;
}

export interface MapInfoOptions {
  /** Selection namespaces (required) */
  readonly selectionNamespaces: string;
  /** Schema definitions */
  readonly schemas: readonly SchemaOptions[];
  /** XML maps */
  readonly maps: readonly MapOptions[];
}

export interface XmlPrOptions {
  /** XML element name */
  readonly xmlElement?: string;
  /** Associated map ID (required) */
  readonly mapId: number;
  /** XPath expression (required) */
  readonly xpath: string;
  /** XML data type (required) */
  readonly xmlDataType: string;
}

export interface XmlCellPrOptions {
  /** Cell ID (required) */
  readonly id: number;
  /** Unique name */
  readonly uniqueName?: string;
  /** XML properties */
  readonly xmlPr: XmlPrOptions;
}

export interface SingleXmlCellOptions {
  /** XML cell ID (required) */
  readonly id: number;
  /** Cell reference, e.g. "A1" (required) */
  readonly r: string;
  /** Connection ID (required) */
  readonly connectionId: number;
  /** XML cell properties */
  readonly xmlCellPr: XmlCellPrOptions;
}

export interface XmlColumnPrOptions {
  /** XPath expression (required) */
  readonly xpath: string;
  /** XML data type (required) */
  readonly xmlDataType: string;
  /** Associated map ID (required) */
  readonly mapId: number;
}

// ── Helper functions ──

function schemaToXml(s: SchemaOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    ID: s.id,
  };
  if (s.schemaRef !== undefined) a.SchemaRef = s.schemaRef;
  if (s.namespace !== undefined) a.Namespace = s.namespace;
  if (s.schemaLanguage !== undefined) a.SchemaLanguage = s.schemaLanguage;
  if (s.schemaID !== undefined) a.SchemaID = s.schemaID;
  if (s.elementFormDefault !== undefined) a.ElementFormDefault = s.elementFormDefault;
  if (s.attributeFormDefault !== undefined) a.AttributeFormDefault = s.attributeFormDefault;
  return `<Schema${attrs(a)}/>`;
}

function dataBindingToXml(db: DataBindingOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {};
  if (db.dataBindingName !== undefined) a.DataBindingName = db.dataBindingName;
  if (db.fileBinding !== undefined) a.FileBinding = db.fileBinding ? 1 : 0;
  if (db.fileBindingName !== undefined) a.FileBindingName = db.fileBindingName;
  if (db.connectionID !== undefined) a.ConnectionID = db.connectionID;
  if (db.dataBindingLoadMode !== undefined) a.DataBindingLoadMode = db.dataBindingLoadMode;
  return `<DataBinding${attrs(a)}/>`;
}

function mapToXml(m: MapOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    ID: m.id,
    Name: m.name,
    RootElement: m.rootElement,
    SchemaID: m.schemaID,
  };
  if (m.showImportExportValidationErrors !== undefined)
    a.ShowImportExportValidationErrors = m.showImportExportValidationErrors ? 1 : 0;
  if (m.append !== undefined) a.Append = m.append ? 1 : 0;
  if (m.dataBindingLoadMode !== undefined) a.DataBindingLoadMode = m.dataBindingLoadMode;
  if (m.autoFit !== undefined) a.AutoFit = m.autoFit ? 1 : 0;
  if (m.fileBinding !== undefined) a.FileBinding = m.fileBinding ? 1 : 0;
  if (m.fileBindingName !== undefined) a.FileBindingName = m.fileBindingName;
  if (m.preserveFormat !== undefined) a.PreserveFormat = m.preserveFormat ? 1 : 0;
  if (m.preserveSortAFLayout !== undefined) a.PreserveSortAFLayout = m.preserveSortAFLayout ? 1 : 0;

  const children: string[] = [];
  if (m.dataBinding) {
    children.push(dataBindingToXml(m.dataBinding));
  }

  if (children.length > 0) {
    return `<Map${attrs(a)}>${children.join("")}</Map>`;
  }
  return `<Map${attrs(a)}/>`;
}

function xmlPrToXml(xp: XmlPrOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    mapId: xp.mapId,
    xpath: xp.xpath,
    xmlDataType: xp.xmlDataType,
  };
  if (xp.xmlElement !== undefined) a.xmlElement = xp.xmlElement;
  return `<xmlPr${attrs(a)}/>`;
}

function xmlCellPrToXml(xcp: XmlCellPrOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    id: xcp.id,
  };
  if (xcp.uniqueName !== undefined) a.uniqueName = xcp.uniqueName;
  return `<xmlCellPr${attrs(a)}>${xmlPrToXml(xcp.xmlPr)}</xmlCellPr>`;
}

function singleXmlCellToXml(sxc: SingleXmlCellOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    id: sxc.id,
    r: sxc.r,
    connectionId: sxc.connectionId,
  };
  return `<singleXmlCell${attrs(a)}>${xmlCellPrToXml(sxc.xmlCellPr)}</singleXmlCell>`;
}

// ── Components ──

export class MapInfoXml extends BaseXmlComponent {
  private readonly options: MapInfoOptions;

  public constructor(options: MapInfoOptions) {
    super("MapInfo");
    this.options = options;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<MapInfo xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    ];
    p.push(` SelectionNamespaces="${this.options.selectionNamespaces}">`);

    for (const s of this.options.schemas) {
      p.push(schemaToXml(s));
    }
    for (const m of this.options.maps) {
      p.push(mapToXml(m));
    }

    p.push("</MapInfo>");
    return p.join("");
  }
}

export class SingleXmlCellsXml extends BaseXmlComponent {
  private readonly cells: readonly SingleXmlCellOptions[];

  public constructor(cells: readonly SingleXmlCellOptions[]) {
    super("singleXmlCells");
    this.cells = cells;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<singleXmlCells xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];

    for (const cell of this.cells) {
      p.push(singleXmlCellToXml(cell));
    }

    p.push("</singleXmlCells>");
    return p.join("");
  }
}

export class XmlColumnPrXml extends BaseXmlComponent {
  private readonly options: XmlColumnPrOptions;

  public constructor(options: XmlColumnPrOptions) {
    super("xmlColumnPr");
    this.options = options;
  }

  public override toXml(_context: Context): string {
    const a: Record<string, string | number | boolean | undefined> = {
      xpath: this.options.xpath,
      xmlDataType: this.options.xmlDataType,
      mapId: this.options.mapId,
    };
    return `<xmlColumnPr${attrs(a)}/>`;
  }
}
