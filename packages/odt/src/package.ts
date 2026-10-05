import type { DataType } from "@office-open/core";

/** Source-only OCF package member carried outside the canonical text model. */
export interface OdtPackageMemberOptions {
  /** Package path below the mimetype root, e.g. `metadata/manifest.rdf`. */
  path: string;
  /** Declared OCF media type, e.g. `application/rdf+xml`. */
  mediaType?: string;
  /** Raw member content: XML text or binary input (base64 for JSON). */
  data: DataType;
}

/** ODF text anchor placement for a draw:frame. */
export type OdtFrameAnchorType = "as-char" | "char" | "page" | "paragraph" | "frame";

/** ODT-specific draw:frame data carried beside the canonical wpsShape. */
export interface OdtEmbeddedObjectFrameOptions {
  /** Canonical `wpsShape.nonVisualProperties.id` that uniquely owns this frame. */
  shapeId: number;
  /** draw:frame anchor placement; omit for the ODF default. */
  anchorType?: OdtFrameAnchorType;
  /** draw:style-name reference into automatic styles. */
  styleName?: string;
}

/** Standalone embedded ODF object with no canonical docx run equivalent. */
export interface OdtEmbeddedObjectOptions {
  /** Object name without its trailing slash, e.g. `Object 1`. */
  path: string;
  /** Declared object media type, e.g. `application/vnd.oasis.opendocument.formula`. */
  mediaType: string;
  /** Members retained verbatim below `path`; their paths are package-absolute. */
  members: OdtPackageMemberOptions[];
  /** Body frames referencing this object; absent for package-only objects. */
  frames?: OdtEmbeddedObjectFrameOptions[];
}
