import type { ChartSpaceOptions } from "@office-open/core/chart";
/**
 * ChartRun types for WordprocessingML documents.
 *
 * Charts are stored as independent XML parts (word/charts/chart{n}.xml)
 * and referenced from document.xml via drawing relationships.
 *
 * @module
 */
import type { Floating } from "@parts/drawing";
import type { GraphicFrameLocksOptions } from "@parts/drawing/descriptor";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { MediaTransformation } from "@shared/media";

/**
 * Options for creating a ChartRun.
 *
 * Extends `ChartSpaceOptions` from `@office-open/core` with DOCX-specific
 * positioning and layout properties.
 *
 * @publicApi
 */
export interface ChartOptions extends ChartSpaceOptions {
  /** Display dimensions */
  transformation: MediaTransformation;
  /** Floating positioning */
  floating?: Floating;
  /** Alternative text for accessibility */
  altText?: DocPropertiesOptions;
  /** Graphic frame locks (wp:cNvGraphicFramePr). `{}` → empty element; omit → authoring default. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Run properties of the wrapping w:r (w:rPr before w:drawing). */
  runProperties?: RunPropertiesOptions;
}
