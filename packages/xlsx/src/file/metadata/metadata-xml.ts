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
import { attrs, escapeXml } from "@office-open/xml";

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

export interface MetadataRecordOptions {
  /** Metadata type index (required) */
  readonly t: number;
  /** Metadata value index (required) */
  readonly v: number;
}

export interface MetadataBlockOptions {
  /** Metadata records */
  readonly records?: readonly MetadataRecordOptions[];
}

export interface MetadataOptions {
  /** Metadata types */
  readonly types?: readonly MetadataTypeOptions[];
  /** Metadata strings */
  readonly strings?: readonly MetadataStringOptions[];
  /** Future metadata */
  readonly futureMetadata?: readonly FutureMetadataOptions[];
  /** MDX metadata (CT_MdxMetadata) */
  readonly mdxMetadata?: readonly MdxOptions[];
  /** Cell metadata blocks */
  readonly cellMetadataBlocks?: readonly MetadataBlockOptions[];
  /** Value metadata blocks */
  readonly valueMetadataBlocks?: readonly MetadataBlockOptions[];
}

/** MDX query entry (CT_Mdx) */
export interface MdxOptions {
  /** MDX function type: "m"|"v"|"s"|"c"|"r"|"p"|"k" */
  readonly f: string;
  /** Name index (required) */
  readonly n: number;
  /** MDX tuple */
  readonly tuple?: MdxTupleOptions;
  /** MDX set */
  readonly set?: MdxSetOptions;
  /** MDX member property */
  readonly memberProp?: MdxMemberPropOptions;
  /** MDX KPI */
  readonly kpi?: MdxKpiOptions;
}

export interface MdxTupleOptions {
  /** Tuple count */
  readonly c?: number;
}

export interface MdxSetOptions {
  /** Namespace count (required) */
  readonly ns: number;
}

export interface MdxMemberPropOptions {
  /** Name index (required) */
  readonly n: number;
  /** Name pair index (required) */
  readonly np: number;
}

export interface MdxKpiOptions {
  /** Name index (required) */
  readonly n: number;
  /** Name pair index (required) */
  readonly np: number;
  /** KPI property (required) */
  readonly p: string;
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

    // cellMetadata
    const cmBlocks = this.opts.cellMetadataBlocks ?? [];
    if (cmBlocks.length > 0) {
      p.push(`<cellMetadata count="${cmBlocks.length}">`);
      for (const blk of cmBlocks) {
        const records = blk.records ?? [];
        if (records.length > 0) {
          const rcParts = records.map((r) => `<rc t="${r.t}" v="${r.v}"/>`);
          p.push(`<bk>${rcParts.join("")}</bk>`);
        } else {
          p.push("<bk/>");
        }
      }
      p.push("</cellMetadata>");
    } else {
      p.push('<cellMetadata count="0"/>');
    }

    // valueMetadata
    const vmBlocks = this.opts.valueMetadataBlocks ?? [];
    if (vmBlocks.length > 0) {
      p.push(`<valueMetadata count="${vmBlocks.length}">`);
      for (const blk of vmBlocks) {
        const records = blk.records ?? [];
        if (records.length > 0) {
          const rcParts = records.map((r) => `<rc t="${r.t}" v="${r.v}"/>`);
          p.push(`<bk>${rcParts.join("")}</bk>`);
        } else {
          p.push("<bk/>");
        }
      }
      p.push("</valueMetadata>");
    } else {
      p.push('<valueMetadata count="0"/>');
    }

    // mdxMetadata (CT_MdxMetadata)
    const mdxMeta = this.opts.mdxMetadata;
    if (mdxMeta && mdxMeta.length > 0) {
      p.push(`<mdxMetadata count="${mdxMeta.length}">`);
      for (const m of mdxMeta) {
        const mAttrs: Record<string, string | number | undefined> = {
          n: m.n,
          f: m.f,
        };
        let child = "";
        if (m.tuple) {
          const tAttrs: Record<string, string | number | undefined> = {};
          if (m.tuple.c !== undefined) tAttrs.c = m.tuple.c;
          child = `<t${attrs(tAttrs)}/>`;
        } else if (m.set) {
          child = `<ms ns="${m.set.ns}"/>`;
        } else if (m.memberProp) {
          child = `<p n="${m.memberProp.n}" np="${m.memberProp.np}"/>`;
        } else if (m.kpi) {
          child = `<k n="${m.kpi.n}" np="${m.kpi.np}" p="${escapeXml(m.kpi.p)}"/>`;
        }
        if (child) {
          p.push(`<mdx${attrs(mAttrs)}>${child}</mdx>`);
        } else {
          p.push(`<mdx${attrs(mAttrs)}/>`);
        }
      }
      p.push("</mdxMetadata>");
    }

    p.push("</metadata>");
    return p.join("");
  }
}
