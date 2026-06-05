/**
 * External Link XML generator — produces xl/externalLinks/externalLinkN.xml.
 *
 * Reference: OOXML transitional, sml.xsd, CT_ExternalLink
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

// ── Options ──

export interface ExternalDefinedNameOptions {
  readonly name: string;
  readonly refersTo?: string;
  readonly sheetId?: number;
}

export interface ExternalCellOptions {
  /** Cell reference, e.g. "A1" */
  readonly reference: string;
  /** Cell data type */
  readonly type?: string;
  /** Cell value */
  readonly value?: string;
}

export interface ExternalBookOptions {
  /** Target path of the external workbook */
  readonly target?: string;
  /** Sheet names from the external workbook */
  readonly sheetNames?: readonly string[];
  /** Defined names from the external workbook */
  readonly definedNames?: readonly ExternalDefinedNameOptions[];
  /** Cached sheet data from the external workbook */
  readonly sheetDataSet?: readonly ExternalSheetDataOptions[];
}

export interface ExternalRowOptions {
  /** Row number (1-based) */
  readonly rowNumber: number;
  readonly cells?: readonly ExternalCellOptions[];
}

export interface ExternalSheetDataOptions {
  readonly sheetId: number;
  readonly refreshError?: boolean;
  readonly rows?: readonly ExternalRowOptions[];
}

export interface ExternalLinkOptions {
  /** External book configuration */
  readonly externalBook?: ExternalBookOptions;
  /** Relationship ID for the external book (set by compiler) */
  readonly bookRId?: string;
  /** OLE link configuration (CT_OleLink) */
  readonly oleLink?: OleLinkOptions;
  /** Relationship ID for the OLE link (set by compiler) */
  readonly oleRId?: string;
}

export interface OleItemOptions {
  /** OLE item name (required) */
  readonly name: string;
  /** Whether to advise events */
  readonly advise?: boolean;
  /** Whether preferred */
  readonly prefer?: boolean;
}

export interface OleLinkOptions {
  /** OLE items */
  readonly oleItems?: readonly OleItemOptions[];
}

// ── Component ──

export class ExternalLinkXml extends BaseXmlComponent {
  private readonly opts: ExternalLinkOptions;

  public constructor(options: ExternalLinkOptions) {
    super("externalLink");
    this.opts = options;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    if (this.opts.externalBook) {
      const book = this.opts.externalBook;
      // r:id is required on externalBook but will be set when the relationship is created
      // We use a placeholder that gets resolved by the compiler
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
          const dnAttrs: Record<string, string | number | undefined> = { name: dn.name };
          if (dn.refersTo !== undefined) dnAttrs.refersTo = dn.refersTo;
          if (dn.sheetId !== undefined) dnAttrs.sheetId = dn.sheetId;
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

      const ridAttr = this.opts.bookRId ? ` r:id="${this.opts.bookRId}"` : "";
      p.push(
        `<externalBook${ridAttr}${bookParts.length > 0 ? `>${bookParts.join("")}</externalBook>` : "/>"}`,
      );
    }

    // oleLink (CT_OleLink)
    if (this.opts.oleLink) {
      const oleRId = this.opts.oleRId ? ` r:id="${escapeXml(this.opts.oleRId)}"` : "";
      const oleChildren: string[] = [];
      if (this.opts.oleLink.oleItems && this.opts.oleLink.oleItems.length > 0) {
        const itemParts: string[] = [`<oleItems>`];
        for (const item of this.opts.oleLink.oleItems) {
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
  }
}
