/**
 * Metadata part types and descriptor for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Metadata
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";
import { attr, attrNum, escapeXml, findChild, stringifyElement } from "@office-open/xml";

// ── Options ──

/** Metadata type declaration (CT_MetadataType). */
export interface MetadataTypeOptions {
  /** Type name (required) */
  name: string;
  /** Minimum supported version (required by XSD) */
  minSupportedVersion: number;
  /** Ghost row flag */
  ghostRow?: boolean;
  /** Ghost column flag */
  ghostCol?: boolean;
  /** Edit flag */
  edit?: boolean;
  /** Delete flag */
  delete?: boolean;
  /** Copy flag */
  copy?: boolean;
  /** Paste all */
  pasteAll?: boolean;
  /** Paste formulas */
  pasteFormulas?: boolean;
  /** Paste values */
  pasteValues?: boolean;
  /** Paste formats */
  pasteFormats?: boolean;
  /** Paste comments */
  pasteComments?: boolean;
  /** Paste data validation */
  pasteDataValidation?: boolean;
  /** Paste borders */
  pasteBorders?: boolean;
  /** Paste column widths */
  pasteColWidths?: boolean;
  /** Paste number formats */
  pasteNumberFormats?: boolean;
  /** Merge cells */
  merge?: boolean;
  /** Split first */
  splitFirst?: boolean;
  /** Split all */
  splitAll?: boolean;
  /** Row/column shift */
  rowColShift?: boolean;
  /** Clear all */
  clearAll?: boolean;
  /** Clear formats */
  clearFormats?: boolean;
  /** Clear contents */
  clearContents?: boolean;
  /** Clear comments */
  clearComments?: boolean;
  /** Assign */
  assign?: boolean;
  /** Coerce */
  coerce?: boolean;
  /** Adjust */
  adjust?: boolean;
  /** Cell metadata */
  cellMeta?: boolean;
}

/** Metadata string entry (CT_MetadataStrings → s, CT_XStringElement @v). */
export interface MetadataStringOptions {
  /** String value (required) */
  value: string;
}

/** Metadata string index reference (CT_MetadataStringIndex). */
export interface MetadataStringIndexOptions {
  /** Index into metadataStrings (@x, required) */
  index: number;
  /** Show string flag */
  show?: boolean;
}

/** MDX tuple (CT_MdxTuple, the t child of mdx). */
export interface MdxTupleOptions {
  /** Tuple count */
  count?: number;
  /** Culture (CT_MdxTuple @ct) */
  culture?: string;
  /** Style index (CT_MdxTuple @si) */
  styleIndex?: number;
  /** Format index (CT_MdxTuple @fi) */
  formatIndex?: number;
  /** Background color ARGB hex (CT_MdxTuple @bc) */
  backgroundColor?: string;
  /** Foreground color ARGB hex (CT_MdxTuple @fc) */
  foregroundColor?: string;
  /** Italic */
  italic?: boolean;
  /** Underline */
  underline?: boolean;
  /** Strikethrough */
  strikethrough?: boolean;
  /** Bold */
  bold?: boolean;
  /** Member name string indexes (n children) */
  stringIndexes?: MetadataStringIndexOptions[];
}

/** MDX set (CT_MdxSet, the ms child of mdx). */
export interface MdxSetOptions {
  /** Namespace count (required) */
  namespaceCount: number;
  /** Tuple count */
  count?: number;
  /** Sort order: "u" | "a" | "d" | "aa" | "ad" | "na" | "nd" (default "u") */
  order?: string;
  /** Member name string indexes (n children) */
  stringIndexes?: MetadataStringIndexOptions[];
}

/** ST_MdxFunctionType — the MDX function discriminator on mdx @f. */
export type MdxFunctionType = "m" | "v" | "s" | "c" | "r" | "p" | "k";

/** ST_MdxKPIProperty — the KPI property kind on k @p. */
export type MdxKpiProperty = "v" | "g" | "s" | "t" | "w" | "m";

/** MDX member property (CT_MdxMemeberProp, the p child of mdx). */
export interface MdxMemberPropOptions {
  /** Index into metadataStrings for the member name (@n, required) */
  nameIndex: number;
  /** Name pair index (@np, required) */
  namePairIndex: number;
}

/** MDX KPI (CT_MdxKPI, the k child of mdx). */
export interface MdxKpiOptions {
  /** Index into metadataStrings for the member name (@n, required) */
  nameIndex: number;
  /** Name pair index (@np, required) */
  namePairIndex: number;
  /** KPI property kind (@p, required) */
  property: MdxKpiProperty;
}

/** MDX metadata entry (CT_Mdx). */
export interface MdxOptions {
  /** MDX function type (@f, required) */
  functionType: MdxFunctionType;
  /** Metadata string index (@n, required) */
  stringIndex: number;
  /** MDX tuple (t) */
  tuple?: MdxTupleOptions;
  /** MDX set (ms) */
  set?: MdxSetOptions;
  /** MDX member property (p) */
  memberProp?: MdxMemberPropOptions;
  /** MDX KPI (k) */
  kpi?: MdxKpiOptions;
}

/** Future metadata block (CT_FutureMetadataBlock). */
export interface FutureMetadataBlockOptions {
  /** Raw extension list preserved verbatim (extLst) */
  extLst?: string;
}

/** Future metadata (CT_FutureMetadata). */
export interface FutureMetadataOptions {
  /** Type name (required) */
  name: string;
  /** Blocks (bk) */
  blocks?: FutureMetadataBlockOptions[];
}

/** Metadata record (CT_MetadataRecord, the rc child of bk). */
export interface MetadataRecordOptions {
  /** Metadata type index (@t, required) */
  typeIndex: number;
  /** Metadata value index (@v, required) */
  valueIndex: number;
}

/** Metadata block (CT_MetadataBlock). */
export interface MetadataBlockOptions {
  /** Records (rc) */
  records?: MetadataRecordOptions[];
}

/** Options for xl/metadata.xml (CT_Metadata). */
export interface MetadataOptions {
  /** Metadata types */
  types?: MetadataTypeOptions[];
  /** Metadata strings */
  strings?: MetadataStringOptions[];
  /** MDX metadata */
  mdx?: MdxOptions[];
  /** Future metadata entries */
  futureMetadata?: FutureMetadataOptions[];
  /** Cell metadata blocks (cellMetadata) */
  cellMetadata?: MetadataBlockOptions[];
  /** Value metadata blocks (valueMetadata) */
  valueMetadata?: MetadataBlockOptions[];
}

// ── Descriptor ──

const METADATA_TYPE_BOOL_ATTRS = [
  "ghostRow",
  "ghostCol",
  "edit",
  "delete",
  "copy",
  "pasteAll",
  "pasteFormulas",
  "pasteValues",
  "pasteFormats",
  "pasteComments",
  "pasteDataValidation",
  "pasteBorders",
  "pasteColWidths",
  "pasteNumberFormats",
  "merge",
  "splitFirst",
  "splitAll",
  "rowColShift",
  "clearAll",
  "clearFormats",
  "clearContents",
  "clearComments",
  "assign",
  "coerce",
  "adjust",
  "cellMeta",
] as const;

export const metadataDesc: CustomDescriptor<MetadataOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];

    if (opts.types && opts.types.length > 0) {
      p.push(`<metadataTypes count="${opts.types.length}">`);
      for (const t of opts.types) {
        const attrs: string[] = [
          `name="${escapeXml(t.name)}"`,
          `minSupportedVersion="${t.minSupportedVersion}"`,
        ];
        for (const key of METADATA_TYPE_BOOL_ATTRS) {
          if (t[key]) attrs.push(`${key}="1"`);
        }
        p.push(`<metadataType ${attrs.join(" ")}/>`);
      }
      p.push("</metadataTypes>");
    }

    if (opts.strings && opts.strings.length > 0) {
      p.push(`<metadataStrings count="${opts.strings.length}">`);
      for (const s of opts.strings) p.push(`<s v="${escapeXml(s.value)}"/>`);
      p.push("</metadataStrings>");
    }

    if (opts.mdx && opts.mdx.length > 0) {
      p.push(`<mdxMetadata count="${opts.mdx.length}">`);
      for (const m of opts.mdx) {
        const attrs: string[] = [`n="${m.stringIndex}"`, `f="${m.functionType}"`];
        let inner = "";
        if (m.tuple) {
          const t = m.tuple;
          const tAttrs: string[] = [];
          if (t.count !== undefined) tAttrs.push(`c="${t.count}"`);
          if (t.culture !== undefined) tAttrs.push(`ct="${escapeXml(t.culture)}"`);
          if (t.styleIndex !== undefined) tAttrs.push(`si="${t.styleIndex}"`);
          if (t.formatIndex !== undefined) tAttrs.push(`fi="${t.formatIndex}"`);
          if (t.backgroundColor !== undefined) tAttrs.push(`bc="${t.backgroundColor}"`);
          if (t.foregroundColor !== undefined) tAttrs.push(`fc="${t.foregroundColor}"`);
          if (t.italic) tAttrs.push('i="1"');
          if (t.underline) tAttrs.push('u="1"');
          if (t.strikethrough) tAttrs.push('st="1"');
          if (t.bold) tAttrs.push('b="1"');
          const idxXml = stringifyStringIndexes(t.stringIndexes);
          inner = idxXml ? `<t ${tAttrs.join(" ")}>${idxXml}</t>` : `<t ${tAttrs.join(" ")}/>`;
        } else if (m.set) {
          const s = m.set;
          const sAttrs: string[] = [`ns="${s.namespaceCount}"`];
          if (s.count !== undefined) sAttrs.push(`c="${s.count}"`);
          if (s.order !== undefined && s.order !== "u") sAttrs.push(`o="${escapeXml(s.order)}"`);
          const idxXml = stringifyStringIndexes(s.stringIndexes);
          inner = idxXml ? `<ms ${sAttrs.join(" ")}>${idxXml}</ms>` : `<ms ${sAttrs.join(" ")}/>`;
        } else if (m.memberProp) {
          inner = `<p n="${m.memberProp.nameIndex}" np="${m.memberProp.namePairIndex}"/>`;
        } else if (m.kpi) {
          inner = `<k n="${m.kpi.nameIndex}" np="${m.kpi.namePairIndex}" p="${m.kpi.property}"/>`;
        }
        p.push(`<mdx ${attrs.join(" ")}>${inner}</mdx>`);
      }
      p.push("</mdxMetadata>");
    }

    for (const fm of opts.futureMetadata ?? []) {
      const blocks = fm.blocks ?? [];
      if (blocks.length > 0) {
        p.push(`<futureMetadata name="${escapeXml(fm.name)}" count="${blocks.length}">`);
        for (const b of blocks) {
          if (b.extLst) p.push(`<bk>${b.extLst}</bk>`);
          else p.push("<bk/>");
        }
        p.push("</futureMetadata>");
      } else {
        p.push(`<futureMetadata name="${escapeXml(fm.name)}"/>`);
      }
    }

    if (opts.cellMetadata && opts.cellMetadata.length > 0) {
      p.push(`<cellMetadata count="${opts.cellMetadata.length}">`);
      for (const b of opts.cellMetadata) p.push(stringifyBlock(b));
      p.push("</cellMetadata>");
    }

    if (opts.valueMetadata && opts.valueMetadata.length > 0) {
      p.push(`<valueMetadata count="${opts.valueMetadata.length}">`);
      for (const b of opts.valueMetadata) p.push(stringifyBlock(b));
      p.push("</valueMetadata>");
    }

    p.push("</metadata>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Partial<MetadataOptions> = {};

    const typesEl = findChild(el, "metadataTypes");
    if (typesEl) {
      const types: MetadataTypeOptions[] = [];
      for (const tEl of typesEl.elements ?? []) {
        if (tEl.name !== "metadataType") continue;
        const t: Partial<MetadataTypeOptions> = { name: attr(tEl, "name") ?? "" };
        const msv = attrNum(tEl, "minSupportedVersion");
        if (msv !== undefined) t.minSupportedVersion = msv;
        for (const key of METADATA_TYPE_BOOL_ATTRS) {
          if (parseOnOff(attr(tEl, key))) t[key] = true;
        }
        types.push(t as MetadataTypeOptions);
      }
      result.types = types;
    }

    const stringsEl = findChild(el, "metadataStrings");
    if (stringsEl) {
      const strings: MetadataStringOptions[] = [];
      for (const sEl of stringsEl.elements ?? []) {
        if (sEl.name !== "s") continue;
        strings.push({ value: attr(sEl, "v") ?? "" });
      }
      result.strings = strings;
    }

    const mdxEl = findChild(el, "mdxMetadata");
    if (mdxEl) {
      const mdx: MdxOptions[] = [];
      for (const mEl of mdxEl.elements ?? []) {
        if (mEl.name !== "mdx") continue;
        const m: Partial<MdxOptions> = {
          stringIndex: attrNum(mEl, "n") ?? 0,
          functionType: (attr(mEl, "f") ?? "m") as MdxFunctionType,
        };
        const tEl = findChild(mEl, "t");
        if (tEl) {
          const t: Partial<MdxTupleOptions> = {};
          const c = attrNum(tEl, "c");
          if (c !== undefined) t.count = c;
          if (attr(tEl, "ct") !== undefined) t.culture = attr(tEl, "ct");
          const si = attrNum(tEl, "si");
          if (si !== undefined) t.styleIndex = si;
          const fi = attrNum(tEl, "fi");
          if (fi !== undefined) t.formatIndex = fi;
          if (attr(tEl, "bc") !== undefined) t.backgroundColor = attr(tEl, "bc");
          if (attr(tEl, "fc") !== undefined) t.foregroundColor = attr(tEl, "fc");
          if (parseOnOff(attr(tEl, "i"))) t.italic = true;
          if (parseOnOff(attr(tEl, "u"))) t.underline = true;
          if (parseOnOff(attr(tEl, "st"))) t.strikethrough = true;
          if (parseOnOff(attr(tEl, "b"))) t.bold = true;
          const idx = parseStringIndexes(tEl);
          if (idx) t.stringIndexes = idx;
          m.tuple = t as MdxTupleOptions;
        }
        const msEl = findChild(mEl, "ms");
        if (msEl) {
          const s: Partial<MdxSetOptions> = { namespaceCount: attrNum(msEl, "ns") ?? 0 };
          const c = attrNum(msEl, "c");
          if (c !== undefined) s.count = c;
          if (attr(msEl, "o") !== undefined) s.order = attr(msEl, "o");
          const idx = parseStringIndexes(msEl);
          if (idx) s.stringIndexes = idx;
          m.set = s as MdxSetOptions;
        }
        const pEl = findChild(mEl, "p");
        if (pEl) {
          m.memberProp = {
            nameIndex: attrNum(pEl, "n") ?? 0,
            namePairIndex: attrNum(pEl, "np") ?? 0,
          };
        }
        const kEl = findChild(mEl, "k");
        if (kEl) {
          m.kpi = {
            nameIndex: attrNum(kEl, "n") ?? 0,
            namePairIndex: attrNum(kEl, "np") ?? 0,
            property: (attr(kEl, "p") ?? "v") as MdxKpiProperty,
          };
        }
        mdx.push(m as MdxOptions);
      }
      result.mdx = mdx;
    }

    const future: FutureMetadataOptions[] = [];
    for (const fmEl of el.elements ?? []) {
      if (fmEl.name !== "futureMetadata") continue;
      const fm: FutureMetadataOptions = { name: attr(fmEl, "name") ?? "" };
      const blocks: FutureMetadataBlockOptions[] = [];
      for (const bEl of fmEl.elements ?? []) {
        if (bEl.name !== "bk") continue;
        const extEl = findChild(bEl, "extLst");
        blocks.push({ extLst: extEl ? stringifyElement(extEl) : undefined });
      }
      if (blocks.length > 0) fm.blocks = blocks;
      future.push(fm);
    }
    if (future.length > 0) result.futureMetadata = future;

    const cellEl = findChild(el, "cellMetadata");
    if (cellEl) result.cellMetadata = parseBlocks(cellEl);
    const valueEl = findChild(el, "valueMetadata");
    if (valueEl) result.valueMetadata = parseBlocks(valueEl);

    return result as MetadataOptions;
  },
};

// ── Helpers ──

function stringifyStringIndexes(indexes: MetadataStringIndexOptions[] | undefined): string {
  if (!indexes || indexes.length === 0) return "";
  return indexes.map((i) => `<n x="${i.index}"${i.show ? ' s="1"' : ""}/>`).join("");
}

function parseStringIndexes(parent: Element): MetadataStringIndexOptions[] | undefined {
  const out: MetadataStringIndexOptions[] = [];
  for (const nEl of parent.elements ?? []) {
    if (nEl.name !== "n") continue;
    const idx: MetadataStringIndexOptions = { index: attrNum(nEl, "x") ?? 0 };
    if (parseOnOff(attr(nEl, "s"))) idx.show = true;
    out.push(idx);
  }
  return out.length > 0 ? out : undefined;
}

function stringifyBlock(b: MetadataBlockOptions): string {
  const rc = (b.records ?? []).map((r) => `<rc t="${r.typeIndex}" v="${r.valueIndex}"/>`).join("");
  return `<bk>${rc}</bk>`;
}

function parseBlocks(el: Element): MetadataBlockOptions[] {
  const blocks: MetadataBlockOptions[] = [];
  for (const bEl of el.elements ?? []) {
    if (bEl.name !== "bk") continue;
    const records: MetadataRecordOptions[] = [];
    for (const rcEl of bEl.elements ?? []) {
      if (rcEl.name !== "rc") continue;
      records.push({ typeIndex: attrNum(rcEl, "t") ?? 0, valueIndex: attrNum(rcEl, "v") ?? 0 });
    }
    blocks.push(records.length > 0 ? { records } : {});
  }
  return blocks;
}
