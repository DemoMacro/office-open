import type { DataType } from "@office-open/core";
import type { OcfManifestOptions } from "@office-open/odf";
import type { WorkbookOptions } from "@office-open/xlsx";

export interface OdsDocumentOptions extends WorkbookOptions {
  /** OCF manifest overlay; model-derived entries are appended when missing. */
  packageManifest?: OcfManifestOptions;
  /** Source-only OCF members retained outside modeled package content. */
  packageMembers?: OdsPackageMemberOptions[];
  /** Source automatic-style overlays keyed by stable name. */
  styleOverlays?: OdsStyleOverlay[];
}

/** Source-only OCF package member carried outside the canonical workbook. */
export interface OdsPackageMemberOptions {
  /** Package path below the mimetype root. */
  path: string;
  /** Declared OCF media type. */
  mediaType?: string;
  /** Raw member content: XML text or binary input (base64 for JSON). */
  data: DataType;
}

export interface OdsStyleOverlay {
  /** Stable source style name referenced by table cells, rows, or columns. */
  name: string;
  /** RNG style family. */
  family: string;
  /** RNG-valid property elements not represented by the canonical WorkbookOptions. */
  properties: { name: string; attributes: Record<string, string | number | boolean> }[];
}
