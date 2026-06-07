/**
 * External Link descriptor — produces xl/externalLinks/externalLinkN.xml.
 *
 * Reference: OOXML transitional, sml.xsd, CT_ExternalLink
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml, findChild } from "@office-open/xml";

import type { ExternalLinkOptions } from "../../file/external-link";

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

    // oleLink (CT_OleLink)
    if (opts.oleLink) {
      const oleRId = opts.oleRId ? ` r:id="${escapeXml(opts.oleRId)}"` : "";
      const oleChildren: string[] = [];
      if (opts.oleLink.oleItems && opts.oleLink.oleItems.length > 0) {
        const itemParts: string[] = [`<oleItems>`];
        for (const item of opts.oleLink.oleItems) {
          const itemAttrs: string[] = [`name="${escapeXml(item.name)}"`];
          if (item.advise) itemAttrs.push('advise="1"');
          if (item.prefer) itemAttrs.push('prefer="1"');
          itemParts.push(`<oleItem ${itemAttrs.join(" ")}/>`);
        }
        itemParts.push("</oleItems>");
        oleChildren.push(itemParts.join(""));
      }
      if (oleChildren.length > 0) {
        p.push(`<oleLink${oleRId}>${oleChildren.join("")}</oleLink>`);
      } else {
        p.push(`<oleLink${oleRId}/>`);
      }
    }

    p.push("</externalLink>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    const bookEl = findChild(el, "externalBook");
    if (bookEl) {
      const book: Record<string, unknown> = {};
      if (bookEl.attributes?.["r:id"]) result.bookRId = bookEl.attributes["r:id"];

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

      result.externalBook = book;
    }

    return result as Record<string, unknown>;
  },
};
