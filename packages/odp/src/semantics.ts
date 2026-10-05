import type { DataType } from "@office-open/core";
import type { OcfManifestOptions } from "@office-open/odf";
import type { PresentationOptions } from "@office-open/pptx";

export interface OdpDocumentOptions extends PresentationOptions {
  /** OCF manifest overlay; model-derived entries are appended when missing. */
  packageManifest?: OcfManifestOptions;
  /** Source-only OCF members retained outside modeled package content. */
  packageMembers?: OdpPackageMemberOptions[];
  /** Source automatic-style overlays keyed by stable name. */
  styleOverlays?: OdpStyleOverlay[];
}

/** Source-only OCF package member carried outside the canonical presentation. */
export interface OdpPackageMemberOptions {
  /** Package path below the mimetype root. */
  path: string;
  /** Declared OCF media type. */
  mediaType?: string;
  /** Raw member content: XML text or binary input (base64 for JSON). */
  data: DataType;
}

export interface OdpStyleOverlay {
  /** Stable source style name. */
  name: string;
  /** RNG style family. */
  family: string;
  /** RNG-valid property elements not represented by canonical slide options. */
  properties: { name: string; attributes: Record<string, string | number | boolean> }[];
}

export type { PresentationOptions, SlideOptions } from "@office-open/pptx";
