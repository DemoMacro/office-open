import type { OdfPackageMemberOptions } from "@office-open/odf";

/** ODF text anchor placement for a draw:frame. */
export type OdtFrameAnchorType = "as-char" | "char" | "page" | "paragraph" | "frame";

/** ODT-specific draw:frame data carried beside the canonical wpsShape. */
export interface OdtEmbeddedObjectFrameOptions {
  shapeId: number;
  anchorType?: OdtFrameAnchorType;
  styleName?: string;
}

/** Standalone embedded ODF object with no canonical docx run equivalent. */
export interface OdtEmbeddedObjectOptions {
  path: string;
  mediaType: string;
  members: OdfPackageMemberOptions[];
  frames?: OdtEmbeddedObjectFrameOptions[];
}
