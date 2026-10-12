import type { DataType } from "@office-open/core";

import type { OcfManifestOptions } from "./container/manifest";

/** Style family allowed by ODF automatic/common style containers. */
export type OdfStyleFamily =
  | "paragraph"
  | "text"
  | "graphic"
  | "table"
  | "table-column"
  | "table-row"
  | "table-cell"
  | "section"
  | "page-layout"
  | "drawing-page";

/** Source-only OCF member carried outside modeled canonical content. */
export interface OdfPackageMemberOptions {
  path: string;
  mediaType?: string;
  data: DataType;
}

/** RNG-valid property element retained after canonical fields are mapped. */
export interface OdfStylePropertyOverlay {
  name: string;
  attributes: Record<string, string | number | boolean>;
}

/** Automatic style retained for an RNG-valid property not in canonical options. */
export interface OdfStyleOverlayOptions {
  name: string;
  family: OdfStyleFamily;
  properties: OdfStylePropertyOverlay[];
}

/** Unified OCF and ODF overlay attached to one canonical family model. */
export interface OdfPackageOverlayOptions {
  packageManifest?: OcfManifestOptions;
  packageMembers?: OdfPackageMemberOptions[];
  styleOverlays?: OdfStyleOverlayOptions[];
}
