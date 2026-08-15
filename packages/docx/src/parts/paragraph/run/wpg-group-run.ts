/**
 * WPG group run types for WordprocessingML documents.
 *
 * @module
 */
import type { EffectListOptions, FillOptions } from "@office-open/core/drawing";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { GroupChildMediaData, MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";
import type { GraphicFrameLocksOptions, GroupShapeLocksOptions } from "../../drawing/descriptor";
import type {
  ChildOffset,
  ChildExtent,
} from "../../drawing/inline/graphic/graphic-data/wpg/wpg-group";

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
  /** Child coordinate offset (chOff) */
  childOffset?: ChildOffset;
  /** Child coordinate extent (chExt) */
  childExtent?: ChildExtent;
  /** Group fill */
  fill?: FillOptions;
  /** Group effects */
  effects?: EffectListOptions;
  floating?: Floating;
  altText?: DocPropertiesOptions;
  /** Raw XML of the mc:Fallback (VML equivalent) — carried verbatim so the full mc:AlternateContent round-trips. */
  vmlFallback?: string;
  /** Media referenced by {@link vmlFallback} `{fileName}` placeholders, registered on generate. */
  vmlFallbackMedia?: BackgroundRawMediaOptions[];
  /** mc:Choice Requires attribute (e.g. "wpg") used to regenerate the AlternateContent wrapper. */
  mcChoiceRequires?: string;
  /** Structured run properties of the wrapping w:r (round-trip) — emitted before the drawing. */
  runProperties?: RunPropertiesOptions;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Group shape locks (wpg:cNvGrpSpPr/a:grpSpLocks) for round-trip. */
  groupShapeLocks?: GroupShapeLocksOptions;
}
