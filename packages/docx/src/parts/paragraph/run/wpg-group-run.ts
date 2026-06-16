/**
 * WPG group run types for WordprocessingML documents.
 *
 * @module
 */
import type { FillOptions } from "@office-open/core/drawingml";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { GroupChildMediaData, MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";
import type { GraphicFrameLocksOptions, GroupShapeLocksOptions } from "../../drawing/descriptor";
import type { EffectListOptions } from "../../drawing/inline/graphic/graphic-data/pic/effects/effect-list";
import type {
  ChildOffset,
  ChildExtent,
} from "../../drawing/inline/graphic/graphic-data/wpg/wpg-group";

export * from "@parts/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreGroupOptions {
  children: GroupChildMediaData[];
  transformation: MediaTransformation;
  /** Child coordinate offset (chOff) */
  chOff?: ChildOffset;
  /** Child coordinate extent (chExt) */
  chExt?: ChildExtent;
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
  /** Raw XML of the wrapping w:r's rPr (round-trip) — emitted before the drawing. */
  runPropertiesRawXml?: string;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Group shape locks (wpg:cNvGrpSpPr/a:grpSpLocks) for round-trip. */
  grpSpLocks?: GroupShapeLocksOptions;
}

/**
 * @publicApi
 */
export type WpgGroupRunOptions = CoreGroupOptions;
