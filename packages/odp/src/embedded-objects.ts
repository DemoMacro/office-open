import type { OdfPackageMemberOptions } from "@office-open/odf";

/** ODP-specific draw:frame data carried beside the canonical pptx shape. */
export interface OdpEmbeddedObjectFrameOptions {
  shapeId: number;
  styleName?: string;
  replacementPath?: string;
}

/** Standalone or linked ODF object with no canonical pptx OLE equivalent. */
export interface OdpEmbeddedObjectOptions {
  path?: string;
  sourceUrl?: string;
  mediaType?: string;
  members: OdfPackageMemberOptions[];
  frames?: OdpEmbeddedObjectFrameOptions[];
}
