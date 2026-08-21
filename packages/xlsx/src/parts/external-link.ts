/**
 * External Link types and descriptor for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_ExternalLink
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml, findChild } from "@office-open/xml";

// ── Types ──

export interface ExternalDefinedNameOptions {
  name: string;
  refersTo?: string;
  sheetId?: number;
  /** Publish to server (CT_DefinedName `@publishToServer`) */
  publishToServer?: boolean;
  /** VBA procedure (CT_DefinedName `@vbProcedure`) */
  vbProcedure?: boolean;
  /** Workbook parameter (CT_DefinedName `@workbookParameter`) */
  workbookParameter?: boolean;
  /** XLM macro (CT_DefinedName `@xlm`) */
  xlm?: boolean;
}

export interface ExternalCellOptions {
  /** Cell reference, e.g. "A1" */
  reference: string;
  /** Cell data type */
  type?: string;
  /** Cell value */
  value?: string;
  /** Value metadata index (CT_ExternalCell `@vm`) */
  valueMetadataIndex?: number;
}

export interface ExternalBookOptions {
  /** Target path of the external workbook */
  target?: string;
  /** Sheet names from the external workbook */
  sheetNames?: string[];
  /** Defined names from the external workbook */
  definedNames?: ExternalDefinedNameOptions[];
  /** Cached sheet data from the external workbook */
  sheetDataSet?: ExternalSheetDataOptions[];
}

export interface ExternalRowOptions {
  /** Row number (1-based) */
  rowNumber: number;
  cells?: ExternalCellOptions[];
}

export interface ExternalSheetDataOptions {
  sheetId: number;
  refreshError?: boolean;
  rows?: ExternalRowOptions[];
}

export interface ExternalLinkOptions {
  /** External book configuration */
  externalBook?: ExternalBookOptions;
  /** Relationship ID for the external book (set by compiler) */
  bookRId?: string;
  /** DDE link configuration (CT_DdeLink) */
  ddeLink?: DdeLinkOptions;
  /** OLE link configuration (CT_OleLink) */
  oleLink?: OleLinkOptions;
  /** Relationship ID for the OLE link (set by compiler) */
  oleRId?: string;
}

/** One cached DDE value (CT_DdeValue — val element + type attribute). */
export interface DdeValueOptions {
  /** Value type (ST_DdeValueType, default "n") */
  type?: "nil" | "b" | "n" | "e" | "str";
  /** Value text (the val element) */
  value: string;
}

/** Cached DDE value grid (CT_DdeValues) */
export interface DdeValuesOptions {
  /** Row count of the cached grid (default 1) */
  rows?: number;
  /** Column count of the cached grid (default 1) */
  cols?: number;
  values: DdeValueOptions[];
}

/** One DDE item (CT_DdeItem) */
export interface DdeItemOptions {
  /** Item name (default "0") */
  name?: string;
  /** Item is an OLE link */
  ole?: boolean;
  /** Advise events on the item */
  advise?: boolean;
  /** Prefer picture representation */
  preferPic?: boolean;
  /** Cached values */
  values?: DdeValuesOptions;
}

/** DDE link (CT_DdeLink — ddeService/ddeTopic + items) */
export interface DdeLinkOptions {
  /** DDE service name */
  ddeService: string;
  ddeTopic: string;
  ddeItems?: DdeItemOptions[];
}

export interface OleItemOptions {
  /** OLE item name (required) */
  name: string;
  /** Show as icon */
  icon?: boolean;
  /** Whether to advise events */
  advise?: boolean;
  /** Prefer picture representation */
  preferPic?: boolean;
}

export interface OleLinkOptions {
  /** OLE program identifier (CT_OleLink `@progId`) */
  progId?: string;
  oleItems?: OleItemOptions[];
}

// ── Descriptor ──

export const externalLinkDesc: CustomDescriptor<ExternalLinkOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    if (opts.externalBook) {
      const book = opts.externalBook;
      const bookParts: string[] = [];

      if (book.sheetNames && book.sheetNames.length > 0) {
        bookParts.push("<sheetNames>");
        for (const name of book.sheetNames) {
          bookParts.push(`<sheetName val="${escapeXml(name)}"/>`);
        }
        bookParts.push("</sheetNames>");
      }

      if (book.definedNames && book.definedNames.length > 0) {
        bookParts.push("<definedNames>");
        for (const dn of book.definedNames) {
          const dnAttrs: Record<string, string | number | boolean | undefined> = { name: dn.name };
          if (dn.refersTo !== undefined) dnAttrs.refersTo = dn.refersTo;
          if (dn.sheetId !== undefined) dnAttrs.sheetId = dn.sheetId;
          if (dn.publishToServer) dnAttrs.publishToServer = 1;
          if (dn.vbProcedure) dnAttrs.vbProcedure = 1;
          if (dn.workbookParameter) dnAttrs.workbookParameter = 1;
          if (dn.xlm) dnAttrs.xlm = 1;
          bookParts.push(`<definedName${attrs(dnAttrs)}/>`);
        }
        bookParts.push("</definedNames>");
      }

      if (book.sheetDataSet && book.sheetDataSet.length > 0) {
        bookParts.push("<sheetDataSet>");
        for (const sd of book.sheetDataSet) {
          const sdAttrs: Record<string, string | number | boolean | undefined> = {
            sheetId: sd.sheetId,
          };
          if (sd.refreshError) sdAttrs.refreshError = 1;
          bookParts.push(`<sheetData${attrs(sdAttrs)}>`);

          if (sd.rows) {
            for (const row of sd.rows) {
              bookParts.push(`<row r="${row.rowNumber}">`);
              if (row.cells) {
                for (const cell of row.cells) {
                  const cellAttrs: Record<string, string | number | undefined> = {
                    r: cell.reference,
                  };
                  if (cell.type !== undefined) cellAttrs.t = cell.type;
                  if (cell.valueMetadataIndex !== undefined) cellAttrs.vm = cell.valueMetadataIndex;
                  if (cell.value !== undefined) {
                    bookParts.push(
                      `<cell${attrs(cellAttrs)}><v>${escapeXml(cell.value)}</v></cell>`,
                    );
                  } else {
                    bookParts.push(`<cell${attrs(cellAttrs)}/>`);
                  }
                }
              }
              bookParts.push("</row>");
            }
          }
          bookParts.push("</sheetData>");
        }
        bookParts.push("</sheetDataSet>");
      }

      const ridAttr = opts.bookRId ? ` r:id="${opts.bookRId}"` : "";
      p.push(
        `<externalBook${ridAttr}${bookParts.length > 0 ? `>${bookParts.join("")}</externalBook>` : "/>"}`,
      );
    }

    // ddeLink (CT_DdeLink — ddeService/ddeTopic + items with cached values)
    if (opts.ddeLink) {
      const dde = opts.ddeLink;
      const ddeParts: string[] = [];
      if (dde.ddeItems && dde.ddeItems.length > 0) {
        ddeParts.push("<ddeItems>");
        for (const item of dde.ddeItems) {
          const itemAttrs: string[] = [];
          if (item.name !== undefined) itemAttrs.push(`name="${escapeXml(item.name)}"`);
          if (item.ole) itemAttrs.push('ole="1"');
          if (item.advise) itemAttrs.push('advise="1"');
          if (item.preferPic) itemAttrs.push('preferPic="1"');
          if (item.values) {
            const v = item.values;
            const valuesAttrs: string[] = [];
            if (v.rows !== undefined) valuesAttrs.push(`rows="${v.rows}"`);
            if (v.cols !== undefined) valuesAttrs.push(`cols="${v.cols}"`);
            const valueParts: string[] = [];
            for (const dv of v.values) {
              const valueAttrs = dv.type !== undefined ? ` t="${dv.type}"` : "";
              valueParts.push(`<value${valueAttrs}><val>${escapeXml(dv.value)}</val></value>`);
            }
            ddeParts.push(
              `<ddeItem${itemAttrs.length > 0 ? " " + itemAttrs.join(" ") : ""}><values${valuesAttrs.length > 0 ? " " + valuesAttrs.join(" ") : ""}>${valueParts.join("")}</values></ddeItem>`,
            );
          } else {
            ddeParts.push(`<ddeItem${itemAttrs.length > 0 ? " " + itemAttrs.join(" ") : ""}/>`);
          }
        }
        ddeParts.push("</ddeItems>");
      }
      p.push(
        `<ddeLink ddeService="${escapeXml(dde.ddeService)}" ddeTopic="${escapeXml(dde.ddeTopic)}"${ddeParts.length > 0 ? `>${ddeParts.join("")}</ddeLink>` : "/>"}`,
      );
    }

    // oleLink (CT_OleLink)
    if (opts.oleLink) {
      const oleRId = opts.oleRId ? ` r:id="${escapeXml(opts.oleRId)}"` : "";
      const progIdAttr = opts.oleLink.progId ? ` progId="${escapeXml(opts.oleLink.progId)}"` : "";
      const oleChildren: string[] = [];
      if (opts.oleLink.oleItems && opts.oleLink.oleItems.length > 0) {
        const itemParts: string[] = [`<oleItems>`];
        for (const item of opts.oleLink.oleItems) {
          const itemAttrs: string[] = [`name="${escapeXml(item.name)}"`];
          if (item.icon) itemAttrs.push('icon="1"');
          if (item.advise) itemAttrs.push('advise="1"');
          if (item.preferPic) itemAttrs.push('preferPic="1"');
          itemParts.push(`<oleItem ${itemAttrs.join(" ")}/>`);
        }
        itemParts.push("</oleItems>");
        oleChildren.push(itemParts.join(""));
      }
      if (oleChildren.length > 0) {
        p.push(`<oleLink${oleRId}${progIdAttr}>${oleChildren.join("")}</oleLink>`);
      } else {
        p.push(`<oleLink${oleRId}${progIdAttr}/>`);
      }
    }

    p.push("</externalLink>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Partial<ExternalLinkOptions> = {};

    const bookEl = findChild(el, "externalBook");
    if (bookEl) {
      const book: ExternalBookOptions = {};
      if (bookEl.attributes?.["r:id"]) result.bookRId = String(bookEl.attributes["r:id"]);

      // sheetNames
      const sheetNamesEl = findChild(bookEl, "sheetNames");
      if (sheetNamesEl) {
        const names: string[] = [];
        for (const child of sheetNamesEl.elements ?? []) {
          if (child.name === "sheetName" && child.attributes?.["val"]) {
            names.push(String(child.attributes["val"]));
          }
        }
        if (names.length > 0) book.sheetNames = names;
      }

      // definedNames
      const definedNamesEl = findChild(bookEl, "definedNames");
      if (definedNamesEl) {
        const dns: ExternalDefinedNameOptions[] = [];
        for (const child of definedNamesEl.elements ?? []) {
          if (child.name !== "definedName") continue;
          const dn: ExternalDefinedNameOptions = { name: String(child.attributes?.["name"] ?? "") };
          if (child.attributes?.["refersTo"]) dn.refersTo = String(child.attributes["refersTo"]);
          if (child.attributes?.["sheetId"] !== undefined)
            dn.sheetId = Number(child.attributes["sheetId"]);
          if (child.attributes?.["publishToServer"]) dn.publishToServer = true;
          if (child.attributes?.["vbProcedure"]) dn.vbProcedure = true;
          if (child.attributes?.["workbookParameter"]) dn.workbookParameter = true;
          if (child.attributes?.["xlm"]) dn.xlm = true;
          dns.push(dn);
        }
        if (dns.length > 0) book.definedNames = dns;
      }

      // sheetDataSet
      const sheetDataSetEl = findChild(bookEl, "sheetDataSet");
      if (sheetDataSetEl) {
        const sds: ExternalSheetDataOptions[] = [];
        for (const sdChild of sheetDataSetEl.elements ?? []) {
          if (sdChild.name !== "sheetData") continue;
          const sd: ExternalSheetDataOptions = {
            sheetId: Number(sdChild.attributes?.["sheetId"] ?? 0),
          };
          if (sdChild.attributes?.["refreshError"]) sd.refreshError = true;

          const rows: ExternalRowOptions[] = [];
          for (const rowChild of sdChild.elements ?? []) {
            if (rowChild.name !== "row") continue;
            const row: ExternalRowOptions = {
              rowNumber: Number(rowChild.attributes?.["r"] ?? 0),
            };

            const cells: ExternalCellOptions[] = [];
            for (const cellChild of rowChild.elements ?? []) {
              if (cellChild.name !== "cell") continue;
              const cell: ExternalCellOptions = {
                reference: String(cellChild.attributes?.["r"] ?? ""),
              };
              if (cellChild.attributes?.["t"]) cell.type = String(cellChild.attributes["t"]);
              if (cellChild.attributes?.["vm"] !== undefined)
                cell.valueMetadataIndex = Number(cellChild.attributes["vm"]);
              const vEl = findChild(cellChild, "v");
              if (vEl && vEl.elements?.[0]?.text !== undefined) {
                cell.value = String(vEl.elements[0].text);
              }
              cells.push(cell);
            }
            if (cells.length > 0) row.cells = cells;
            rows.push(row);
          }
          if (rows.length > 0) sd.rows = rows;
          sds.push(sd);
        }
        if (sds.length > 0) book.sheetDataSet = sds;
      }

      result.externalBook = book;
    }

    // ddeLink
    const ddeEl = findChild(el, "ddeLink");
    if (ddeEl) {
      const dde: DdeLinkOptions = {
        ddeService: String(ddeEl.attributes?.["ddeService"] ?? ""),
        ddeTopic: String(ddeEl.attributes?.["ddeTopic"] ?? ""),
      };

      const ddeItemsEl = findChild(ddeEl, "ddeItems");
      if (ddeItemsEl) {
        const items: DdeItemOptions[] = [];
        for (const child of ddeItemsEl.elements ?? []) {
          if (child.name !== "ddeItem") continue;
          const item: DdeItemOptions = {};
          if (child.attributes?.["name"] !== undefined)
            item.name = String(child.attributes["name"]);
          if (child.attributes?.["ole"]) item.ole = true;
          if (child.attributes?.["advise"]) item.advise = true;
          if (child.attributes?.["preferPic"]) item.preferPic = true;

          const valuesEl = findChild(child, "values");
          if (valuesEl) {
            const values: DdeValuesOptions = { values: [] };
            if (valuesEl.attributes?.["rows"] !== undefined)
              values.rows = Number(valuesEl.attributes["rows"]);
            if (valuesEl.attributes?.["cols"] !== undefined)
              values.cols = Number(valuesEl.attributes["cols"]);
            for (const valueChild of valuesEl.elements ?? []) {
              if (valueChild.name !== "value") continue;
              const valEl = findChild(valueChild, "val");
              const dv: DdeValueOptions = {
                value: valEl ? String(valEl.elements?.[0]?.text ?? "") : "",
              };
              if (valueChild.attributes?.["t"])
                dv.type = valueChild.attributes["t"] as DdeValueOptions["type"];
              values.values.push(dv);
            }
            item.values = values;
          }
          items.push(item);
        }
        if (items.length > 0) dde.ddeItems = items;
      }

      result.ddeLink = dde;
    }

    // oleLink
    const oleEl = findChild(el, "oleLink");
    if (oleEl) {
      const ole: OleLinkOptions = {};
      if (oleEl.attributes?.["r:id"]) result.oleRId = String(oleEl.attributes["r:id"]);
      if (oleEl.attributes?.["progId"]) ole.progId = String(oleEl.attributes["progId"]);

      const oleItemsEl = findChild(oleEl, "oleItems");
      if (oleItemsEl) {
        const items: OleItemOptions[] = [];
        for (const child of oleItemsEl.elements ?? []) {
          if (child.name !== "oleItem") continue;
          const item: OleItemOptions = { name: String(child.attributes?.["name"] ?? "") };
          if (child.attributes?.["icon"]) item.icon = true;
          if (child.attributes?.["advise"]) item.advise = true;
          if (child.attributes?.["preferPic"]) item.preferPic = true;
          items.push(item);
        }
        if (items.length > 0) ole.oleItems = items;
      }

      result.oleLink = ole;
    }

    return result as ExternalLinkOptions;
  },
};
