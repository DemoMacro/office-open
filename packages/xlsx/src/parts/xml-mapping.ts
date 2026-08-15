/**
 * XML Mapping types and descriptors for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd
 * CT_MapInfo, CT_Schema, CT_Map, CT_DataBinding,
 * CT_SingleXmlCells, CT_SingleXmlCell, CT_XmlCellPr, CT_XmlPr
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";
import { attr, attrNum, escapeXml, findChild, stringifyElement } from "@office-open/xml";

// ── Options ──

/** XML schema definition (CT_Schema). Content is arbitrary (xsd:any, mixed). */
export interface SchemaOptions {
  /** Unique schema ID (required) */
  id: string;
  /** Schema reference (CT_Schema @SchemaRef) */
  schemaRef?: string;
  /** Namespace URI (CT_Schema @Namespace) */
  namespace?: string;
  /** Schema language (CT_Schema @SchemaLanguage) */
  schemaLanguage?: string;
  /** Raw schema content preserved verbatim (xsd:any inner XML) */
  content?: string;
}

/** Data binding (CT_DataBinding). */
export interface DataBindingOptions {
  /** Data binding name (CT_DataBinding@DataBindingName) */
  dataBindingName?: string;
  /** Whether this is a file binding (CT_DataBinding @FileBinding) */
  fileBinding?: boolean;
  /** Connection ID (CT_DataBinding @ConnectionID) */
  connectionId?: number;
  /** File binding name (CT_DataBinding @FileBindingName) */
  fileBindingName?: string;
  /** Load mode (CT_DataBinding@DataBindingLoadMode, required by XSD) */
  dataBindingLoadMode?: number;
}

/** XML map (CT_Map). */
export interface MapOptions {
  /** Unique map ID (required) */
  id: number;
  /** Map name (required) */
  name: string;
  /** Root element name (required) */
  rootElement: string;
  /** Owning schema ID (required) */
  schemaId: string;
  /** Show import/export validation errors (required by XSD) */
  showImportExportValidationErrors?: boolean;
  /** Auto-fit columns (CT_Map @AutoFit, required by XSD) */
  autoFit?: boolean;
  /** Append new data (CT_Map @Append, required by XSD) */
  append?: boolean;
  /** Preserve sort/autofilter layout (CT_Map @PreserveSortAFLayout, required by XSD) */
  preserveSortAFLayout?: boolean;
  /** Preserve formatting (CT_Map @PreserveFormat, required by XSD) */
  preserveFormat?: boolean;
  /** Data binding (DataBinding child) */
  dataBinding?: DataBindingOptions;
}

/** Options for xl/xmlMaps.xml (CT_MapInfo). */
export interface MapInfoOptions {
  /** Selection namespaces declaration (required) */
  selectionNamespaces: string;
  /** Schema definitions */
  schemas: SchemaOptions[];
  /** XML maps */
  maps: MapOptions[];
}

/** XML cell properties (CT_XmlPr). */
export interface XmlPropertiesOptions {
  /** Owning map ID (required) */
  mapId: number;
  /** XPath expression (required) */
  xpath: string;
  /** XML data type (required) */
  xmlDataType: string;
}

/** Cell properties of a single-cell XML table (CT_XmlCellPr). */
export interface XmlCellPropertiesOptions {
  /** Table ID (required) */
  id: number;
  /** Unique name (CT_XmlCellPr @uniqueName) */
  uniqueName?: string;
  /** XML properties (required) */
  xmlPr: XmlPropertiesOptions;
}

/** Single-cell XML table entry (CT_SingleXmlCell). */
export interface SingleXmlCellOptions {
  /** Table ID (required) */
  id: number;
  /** Cell reference, e.g. "A1" (required) */
  r: string;
  /** Connection ID (required) */
  connectionId: number;
  /** Cell properties (required) */
  xmlCellPr: XmlCellPropertiesOptions;
}

/** Options for xl/tables/tableSingleCells{n}.xml (CT_SingleXmlCells). */
export interface SingleXmlCellsOptions {
  /** Single-cell XML tables */
  cells: SingleXmlCellOptions[];
}

/** Table column XML mapping (CT_XmlColumnPr) — lives on table columns, not this part. */
export interface XmlColumnPropertiesOptions {
  /** XPath expression (required) */
  xpath: string;
  /** XML data type (required) */
  xmlDataType: string;
  /** Owning map ID (required) */
  mapId: number;
  /** Denormalized (CT_XmlColumnPr @denormalized) */
  denormalized?: boolean;
}

// ── Descriptors ──

export const mapInfoDesc: CustomDescriptor<MapInfoOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      `<MapInfo xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" SelectionNamespaces="${escapeXml(opts.selectionNamespaces)}">`,
    ];
    for (const s of opts.schemas) {
      const attrs: string[] = [`ID="${escapeXml(s.id)}"`];
      if (s.schemaRef !== undefined) attrs.push(`SchemaRef="${escapeXml(s.schemaRef)}"`);
      if (s.namespace !== undefined) attrs.push(`Namespace="${escapeXml(s.namespace)}"`);
      if (s.schemaLanguage !== undefined)
        attrs.push(`SchemaLanguage="${escapeXml(s.schemaLanguage)}"`);
      if (s.content) p.push(`<Schema ${attrs.join(" ")}>${s.content}</Schema>`);
      else p.push(`<Schema ${attrs.join(" ")}/>`);
    }
    for (const m of opts.maps) {
      const attrs: string[] = [
        `ID="${m.id}"`,
        `Name="${escapeXml(m.name)}"`,
        `RootElement="${escapeXml(m.rootElement)}"`,
        `SchemaID="${escapeXml(m.schemaId)}"`,
        `ShowImportExportValidationErrors="${m.showImportExportValidationErrors ? 1 : 0}"`,
        `AutoFit="${m.autoFit ? 1 : 0}"`,
        `Append="${m.append ? 1 : 0}"`,
        `PreserveSortAFLayout="${m.preserveSortAFLayout ? 1 : 0}"`,
        `PreserveFormat="${m.preserveFormat ? 1 : 0}"`,
      ];
      if (m.dataBinding) {
        const d = m.dataBinding;
        const dAttrs: string[] = [];
        if (d.dataBindingName !== undefined)
          dAttrs.push(`DataBindingName="${escapeXml(d.dataBindingName)}"`);
        if (d.fileBinding) dAttrs.push('FileBinding="1"');
        if (d.connectionId !== undefined) dAttrs.push(`ConnectionID="${d.connectionId}"`);
        if (d.fileBindingName !== undefined)
          dAttrs.push(`FileBindingName="${escapeXml(d.fileBindingName)}"`);
        dAttrs.push(`DataBindingLoadMode="${d.dataBindingLoadMode ?? 0}"`);
        p.push(`<Map ${attrs.join(" ")}><DataBinding ${dAttrs.join(" ")}/></Map>`);
      } else {
        p.push(`<Map ${attrs.join(" ")}/>`);
      }
    }
    p.push("</MapInfo>");
    return p.join("");
  },

  parse(el, _ctx) {
    const schemas: SchemaOptions[] = [];
    const maps: MapOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "Schema") {
        const s: Partial<SchemaOptions> = { id: attr(child, "ID") ?? "" };
        if (attr(child, "SchemaRef") !== undefined) s.schemaRef = attr(child, "SchemaRef");
        if (attr(child, "Namespace") !== undefined) s.namespace = attr(child, "Namespace");
        if (attr(child, "SchemaLanguage") !== undefined)
          s.schemaLanguage = attr(child, "SchemaLanguage");
        const inner = (child.elements ?? []).map((c) => stringifyElement(c)).join("");
        if (inner) s.content = inner;
        schemas.push(s as SchemaOptions);
      } else if (child.name === "Map") {
        const m: Partial<MapOptions> = {
          id: attrNum(child, "ID") ?? 0,
          name: attr(child, "Name") ?? "",
          rootElement: attr(child, "RootElement") ?? "",
          schemaId: attr(child, "SchemaID") ?? "",
        };
        if (parseOnOff(attr(child, "ShowImportExportValidationErrors")))
          m.showImportExportValidationErrors = true;
        if (parseOnOff(attr(child, "AutoFit"))) m.autoFit = true;
        if (parseOnOff(attr(child, "Append"))) m.append = true;
        if (parseOnOff(attr(child, "PreserveSortAFLayout"))) m.preserveSortAFLayout = true;
        if (parseOnOff(attr(child, "PreserveFormat"))) m.preserveFormat = true;
        const dEl = findChild(child, "DataBinding");
        if (dEl) {
          const d: DataBindingOptions = {};
          if (attr(dEl, "DataBindingName") !== undefined)
            d.dataBindingName = attr(dEl, "DataBindingName");
          if (parseOnOff(attr(dEl, "FileBinding"))) d.fileBinding = true;
          const cid = attrNum(dEl, "ConnectionID");
          if (cid !== undefined) d.connectionId = cid;
          if (attr(dEl, "FileBindingName") !== undefined)
            d.fileBindingName = attr(dEl, "FileBindingName");
          const mode = attrNum(dEl, "DataBindingLoadMode");
          if (mode !== undefined) d.dataBindingLoadMode = mode;
          m.dataBinding = d;
        }
        maps.push(m as MapOptions);
      }
    }
    return {
      selectionNamespaces: attr(el, "SelectionNamespaces") ?? "",
      schemas,
      maps,
    };
  },
};

export const singleXmlCellsDesc: CustomDescriptor<SingleXmlCellsOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<singleXmlCells xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];
    for (const c of opts.cells) {
      const pr = c.xmlCellPr;
      const prAttrs: string[] = [`id="${pr.id}"`];
      if (pr.uniqueName !== undefined) prAttrs.push(`uniqueName="${escapeXml(pr.uniqueName)}"`);
      p.push(
        `<singleXmlCell id="${c.id}" r="${escapeXml(c.r)}" connectionId="${c.connectionId}">` +
          `<xmlCellPr ${prAttrs.join(" ")}>` +
          `<xmlPr mapId="${pr.xmlPr.mapId}" xpath="${escapeXml(pr.xmlPr.xpath)}" xmlDataType="${escapeXml(pr.xmlPr.xmlDataType)}"/>` +
          `</xmlCellPr></singleXmlCell>`,
      );
    }
    p.push("</singleXmlCells>");
    return p.join("");
  },

  parse(el: Element, _ctx) {
    const cells: SingleXmlCellOptions[] = [];
    for (const cEl of el.elements ?? []) {
      if (cEl.name !== "singleXmlCell") continue;
      const prEl = findChild(cEl, "xmlCellPr");
      if (!prEl) continue;
      const xmlPrEl = findChild(prEl, "xmlPr");
      if (!xmlPrEl) continue;
      const pr: Partial<XmlCellPropertiesOptions> = { id: attrNum(prEl, "id") ?? 0 };
      if (attr(prEl, "uniqueName") !== undefined) pr.uniqueName = attr(prEl, "uniqueName");
      pr.xmlPr = {
        mapId: attrNum(xmlPrEl, "mapId") ?? 0,
        xpath: attr(xmlPrEl, "xpath") ?? "",
        xmlDataType: attr(xmlPrEl, "xmlDataType") ?? "",
      };
      cells.push({
        id: attrNum(cEl, "id") ?? 0,
        r: attr(cEl, "r") ?? "",
        connectionId: attrNum(cEl, "connectionId") ?? 0,
        xmlCellPr: pr as XmlCellPropertiesOptions,
      });
    }
    return { cells };
  },
};
