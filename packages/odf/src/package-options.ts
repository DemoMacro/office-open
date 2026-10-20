import type { DataType } from "@office-open/core";

import type { OcfManifestOptions } from "./container/manifest";

/** Style family allowed by ODF automatic/common style containers. */
export type StyleFamily =
  | "paragraph"
  | "text"
  | "graphic"
  | "table"
  | "table-column"
  | "table-row"
  | "table-cell"
  | "section"
  | "page-layout"
  | "drawing-page"
  | "ruby";

/** Source-only OCF member carried outside modeled canonical content. */
export interface PackageMemberOptions {
  path: string;
  mediaType?: string;
  data: DataType;
}

/** RNG-valid property element retained after canonical fields are mapped. */
export interface StylePropertyOptions {
  name: string;
  attributes: Record<string, string | number | boolean>;
}

/** Automatic style retained for an RNG-valid property not in canonical options. */
export interface StyleResourceOptions {
  name: string;
  family: StyleFamily;
  properties: StylePropertyOptions[];
}

/** Unified OCF and style resources attached to one canonical family model. */
export interface PackageOptions {
  packageManifest?: OcfManifestOptions;
  packageMembers?: PackageMemberOptions[];
  styleResources?: StyleResourceOptions[];
}
