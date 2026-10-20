/**
 * WPG group run types for WordprocessingML documents.
 *
 * @module
 */
import type { LongHexNumber } from "@office-open/core";
import type {
  BlackWhiteMode,
  EffectListOptions,
  FillOptions,
  OfficeArtExtensionOptions,
} from "@office-open/core/drawing";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { GroupChildMediaData, MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";
import type {
  DrawingExtensionIds,
  GraphicFrameLocksOptions,
  GroupShapeLocksOptions,
} from "../../drawing/descriptor";
export * from "@parts/drawing/inline/graphic/graphic-data/wps/body-properties";

/**
 * Group options for docx (wpg:wgp). The cNvPr fields bridge through altText
 * (DocPropertiesOptions); the rest is the docx run-level group model.
 *
 * @publicApi
 */
export interface GroupOptions {
  children: GroupChildMediaData[];
  transformation: MediaTransformation;
  /**
   * Group extent from wpg:grpSpPr/a:xfrm/a:ext in EMUs when it differs from
   * wp:extent; omitted when the two source sizes agree.
   */
  groupExtent?: { width: number; height: number };
  /** Child coordinate system offset X (a:chOff @x, EMU). */
  childOffsetX?: number;
  /** Child coordinate system offset Y (a:chOff @y, EMU). */
  childOffsetY?: number;
  /** Source carried a:chOff; round-trip only. */
  childOffsetPresent?: boolean;
  /** Child coordinate system extent width (a:chExt @cx, EMU). */
  childExtentWidth?: number;
  /** Child coordinate system extent height (a:chExt @cy, EMU). */
  childExtentHeight?: number;
  /** Source carried a:chExt; round-trip only. */
  childExtentPresent?: boolean;
  /** Group fill */
  fill?: FillOptions;
  /** Black-and-white rendering mode on wpg:grpSpPr (`@bwMode`). */
  blackWhiteMode?: BlackWhiteMode;
  /** Group effects */
  effects?: EffectListOptions;
  floating?: Floating;
  altText?: DocPropertiesOptions;
  /**
   * Raw XML of the mc:Fallback (VML equivalent) — carried verbatim so the full
   * mc:AlternateContent round-trips. Round-trip channel: captured from a
   * parsed source document — do not hand-author.
   */
  vmlFallback?: string;
  /** Media referenced by `vmlFallback` `{fileName}` placeholders, registered on generate. */
  vmlFallbackMedia?: BackgroundRawMediaOptions[];
  /** mc:Choice Requires attribute (e.g. "wpg") used to regenerate the AlternateContent wrapper. */
  mcChoiceRequires?: string;
  /** Structured run properties of the wrapping w:r (round-trip) — emitted before the drawing. */
  runProperties?: RunPropertiesOptions;
  /** Revision save ID of the wrapping w:r (w:rsidR, round-trip). */
  additionRsid?: LongHexNumber;
  /** Revision save ID when the wrapping w:r was deleted (w:rsidDel, round-trip). */
  deletionRsid?: LongHexNumber;
  /** Revision save ID of the wrapping w:r properties (w:rsidRPr, round-trip). */
  runPropertiesRsid?: LongHexNumber;
  /** A w:lastRenderedPageBreak shared the drawing's run (round-trip) — emitted before the drawing. */
  lastRenderedPageBreak?: boolean;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Word 2010 wrapper ids; round-trip only. */
  extensionIds?: DrawingExtensionIds;
  /** Group shape locks (wpg:cNvGrpSpPr/a:grpSpLocks) for round-trip. */
  groupShapeLocks?: GroupShapeLocksOptions;
  /** Group property extensions (wpg:grpSpPr/a:extLst). */
  extensions?: OfficeArtExtensionOptions[];
}
