/**
 * Metadata XML generator — produces xl/metadata.xml.
 *
 * Contains cell metadata and value metadata for rich data types.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Metadata
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs } from "@office-open/xml";

// ── Options ──

export interface MetadataTypeOptions {
  /** Metadata type name */
  readonly name: string;
  /** Minimum version */
  readonly minVersion?: number;
  /** Minimum supported version (CT_MetadataType @minSupportedVersion) */
  readonly minSupportedVersion?: number;
  /** Ghost row flag */
  readonly ghostRow?: boolean;
  /** Ghost column flag */
  readonly ghostCol?: boolean;
  /** Edit flag */
  readonly edit?: boolean;
  /** Delete flag */
  readonly delete?: boolean;
  /** Copy flag */
  readonly copy?: boolean;
  /** Paste flag */
  readonly paste?: boolean;
  /** Paste all (CT_MetadataType @pasteAll) */
  readonly pasteAll?: boolean;
  /** Paste formulas */
  readonly pasteFormulas?: boolean;
  /** Paste values */
  readonly pasteValues?: boolean;
  /** Paste formats */
  readonly pasteFormats?: boolean;
  /** Paste comments */
  readonly pasteComments?: boolean;
  /** Paste data validation */
  readonly pasteDataValidation?: boolean;
  /** Paste borders */
  readonly pasteBorders?: boolean;
  /** Paste column widths */
  readonly pasteColWidths?: boolean;
  /** Paste number formats */
  readonly pasteNumberFormats?: boolean;
  /** Merge cells */
  readonly merge?: boolean;
  /** Split first */
  readonly splitFirst?: boolean;
  /** Split all */
  readonly splitAll?: boolean;
  /** Row/column shift */
  readonly rowColShift?: boolean;
  /** Clear all */
  readonly clearAll?: boolean;
  /** Clear formats */
  readonly clearFormats?: boolean;
  /** Clear contents */
  readonly clearContents?: boolean;
  /** Clear comments */
  readonly clearComments?: boolean;
  /** Assign */
  readonly assign?: boolean;
  /** Coerce */
  readonly coerce?: boolean;
  /** Adjust */
  readonly adjust?: boolean;
  /** Cell metadata */
  readonly cellMeta?: boolean;
}

export interface MetadataStringOptions {
  /** String value */
  readonly value: string;
}

export interface FutureMetadataOptions {
  /** Future metadata type name */
  readonly name: string;
  /** Future metadata type */
  readonly type: string;
}

export interface MetadataOptions {
  /** Metadata types */
  readonly types?: readonly MetadataTypeOptions[];
  /** Metadata strings */
  readonly strings?: readonly MetadataStringOptions[];
  /** Future metadata */
  readonly futureMetadata?: readonly FutureMetadataOptions[];
}

// ── Component ──

export class MetadataXml extends BaseXmlComponent {
  private readonly opts: MetadataOptions;

  public constructor(options: MetadataOptions) {
    super("metadata");
    this.opts = options;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];

    // metadataTypes
    const types = this.opts.types ?? [];
    p.push(`<metadataTypes count="${types.length}">`);
    for (const t of types) {
      p.push(
        `<metadataType${attrs({
          name: t.name,
          minVersion: t.minVersion,
          minSupportedVersion: t.minSupportedVersion,
          ghostRow: t.ghostRow ? 1 : undefined,
          ghostCol: t.ghostCol ? 1 : undefined,
          edit: t.edit ? 1 : undefined,
          delete: t.delete ? 1 : undefined,
          copy: t.copy ? 1 : undefined,
          paste: t.paste ? 1 : undefined,
          pasteAll: t.pasteAll ? 1 : undefined,
          pasteFormulas: t.pasteFormulas ? 1 : undefined,
          pasteValues: t.pasteValues ? 1 : undefined,
          pasteFormats: t.pasteFormats ? 1 : undefined,
          pasteComments: t.pasteComments ? 1 : undefined,
          pasteDataValidation: t.pasteDataValidation ? 1 : undefined,
          pasteBorders: t.pasteBorders ? 1 : undefined,
          pasteColWidths: t.pasteColWidths ? 1 : undefined,
          pasteNumberFormats: t.pasteNumberFormats ? 1 : undefined,
          merge: t.merge ? 1 : undefined,
          splitFirst: t.splitFirst ? 1 : undefined,
          splitAll: t.splitAll ? 1 : undefined,
          rowColShift: t.rowColShift ? 1 : undefined,
          clearAll: t.clearAll ? 1 : undefined,
          clearFormats: t.clearFormats ? 1 : undefined,
          clearContents: t.clearContents ? 1 : undefined,
          clearComments: t.clearComments ? 1 : undefined,
          assign: t.assign ? 1 : undefined,
          coerce: t.coerce ? 1 : undefined,
          adjust: t.adjust ? 1 : undefined,
          cellMeta: t.cellMeta ? 1 : undefined,
        })}/>`,
      );
    }
    p.push("</metadataTypes>");

    // metadataStrings
    const strings = this.opts.strings ?? [];
    if (strings.length > 0) {
      p.push(`<metadataStrings count="${strings.length}">`);
      for (const s of strings) {
        p.push(`<s v="${s.value}"/>`);
      }
      p.push("</metadataStrings>");
    }

    // futureMetadata
    const future = this.opts.futureMetadata ?? [];
    if (future.length > 0) {
      for (const f of future) {
        p.push(`<futureMetadata name="${f.name}" type="${f.type}"><bk/>`);
        p.push("</futureMetadata>");
      }
    }

    // cellMetadata & valueMetadata (empty containers)
    p.push('<cellMetadata count="0"/>');
    p.push('<valueMetadata count="0"/>');

    p.push("</metadata>");
    return p.join("");
  }
}
