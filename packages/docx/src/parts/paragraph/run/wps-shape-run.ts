import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
/**
 * WPS shape run types for WordprocessingML documents.
 *
 * @module
 */
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { WpsShapeCoreOptions } from "@parts/drawing/inline/graphic/graphic-data/wps";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";
import type { GraphicFrameLocksOptions } from "../../drawing/descriptor";

export * from "@parts/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreShapeOptions {
  transformation: MediaTransformation;
  floating?: Floating;
  altText?: DocPropertiesOptions;
  /** Raw XML of the mc:Fallback (VML equivalent) — carried verbatim so the full mc:AlternateContent round-trips. */
  vmlFallback?: string;
  /** Media referenced by {@link vmlFallback} `{fileName}` placeholders, registered on generate. */
  vmlFallbackMedia?: BackgroundRawMediaOptions[];
  /** mc:Choice Requires attribute (e.g. "wps") used to regenerate the AlternateContent wrapper. */
  mcChoiceRequires?: string;
  /** Structured run properties of the wrapping w:r (round-trip) — emitted before the drawing. */
  runProperties?: RunPropertiesOptions;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
}

/**
 * @publicApi
 */
export type WpsShapeRunOptions = WpsShapeCoreOptions & CoreShapeOptions;
