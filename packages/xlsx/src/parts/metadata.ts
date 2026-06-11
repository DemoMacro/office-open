/**
 * Metadata options for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Metadata
 *
 * @module
 */

export interface MetadataTypeOptions {
  /** Metadata type name */
  name: string;
  /** Minimum version */
  minVersion?: number;
  /** Minimum supported version (CT_MetadataType @minSupportedVersion) */
  minSupportedVersion?: number;
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
  /** Paste flag */
  paste?: boolean;
  /** Paste all (CT_MetadataType @pasteAll) */
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

export interface MetadataStringOptions {
  /** String value */
  value: string;
}

export interface FutureMetadataOptions {
  /** Future metadata type name */
  name: string;
  /** Future metadata type */
  type: string;
}

export interface MetadataOptions {
  /** Metadata types */
  types?: MetadataTypeOptions[];
  /** Metadata strings */
  strings?: MetadataStringOptions[];
  /** Future metadata */
  futureMetadata?: FutureMetadataOptions[];
  /** MDX metadata (CT_MdxMetadata) */
  mdxMetadata?: MdxOptions[];
  /** Cell metadata blocks */
  cellMetadataBlocks?: MetadataBlockOptions[];
  /** Value metadata blocks */
  valueMetadataBlocks?: MetadataBlockOptions[];
}

/** MDX query entry (CT_Mdx) */
export interface MdxOptions {
  /** MDX function type: "m"|"v"|"s"|"c"|"r"|"p"|"k" */
  f: string;
  /** Name index (required) */
  n: number;
  /** MDX tuple */
  tuple?: MdxTupleOptions;
  /** MDX set */
  set?: MdxSetOptions;
  /** MDX member property */
  memberProp?: MdxMemberPropOptions;
  /** MDX KPI */
  kpi?: MdxKpiOptions;
}

export interface MdxTupleOptions {
  /** Tuple count */
  c?: number;
}

export interface MdxSetOptions {
  /** Namespace count (required) */
  ns: number;
}

export interface MdxMemberPropOptions {
  /** Name index (required) */
  n: number;
  /** Name pair index (required) */
  np: number;
}

export interface MdxKpiOptions {
  /** Name index (required) */
  n: number;
  /** Name pair index (required) */
  np: number;
  /** KPI property (required) */
  p: string;
}

// ── Internal types used during compilation ──

export interface MetadataRecordOptions {
  /** Metadata type index (required) */
  t: number;
  /** Metadata value index (required) */
  v: number;
}

export interface MetadataBlockOptions {
  /** Metadata records */
  records?: MetadataRecordOptions[];
}
