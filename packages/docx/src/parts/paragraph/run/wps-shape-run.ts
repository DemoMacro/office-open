import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
/**
 * Shape run types for WordprocessingML documents.
 *
 * @module
 */
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { ShapeCoreOptions } from "@parts/drawing/inline/graphic/graphic-data/wps";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";
import type { GraphicFrameLocksOptions } from "../../drawing/descriptor";

export * from "@parts/drawing/inline/graphic/graphic-data/wps/body-properties";
export type {
  ShapeCoreOptions,
  ShapeTextBoxChild,
  TextBoxPartOptions,
} from "@parts/drawing/inline/graphic/graphic-data/wps/wps-shape";

interface ShapeRunOptions {
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
  /** A w:lastRenderedPageBreak shared the drawing's run (round-trip) — emitted before the drawing. */
  lastRenderedPageBreak?: boolean;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
}

/**
 * Shape options for docx (wps:wsp). The shape body comes from
 * {@link ShapeCoreOptions}; the rest is the docx run-level shape model
 * (transformation, floating, altText, run wrapping).
 *
 * @publicApi
 */
export type ShapeOptions = ShapeCoreOptions & ShapeRunOptions;
