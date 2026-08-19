/**
 * SmartArtRun types for WordprocessingML documents.
 *
 * SmartArt data is stored in `word/diagrams/data{n}.xml`.
 * Layout, style, and colors reference Word's built-in definitions or carry
 * full custom definitions.
 *
 * @module
 */
import type {
  ColorDefinitionOptions,
  LayoutDefinitionOptions,
  SmartArtRawParts,
  StyleDefinitionOptions,
} from "@office-open/core/smartart";
import type { Floating } from "@parts/drawing";
import type { GraphicFrameLocksOptions } from "@parts/drawing/descriptor";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { MediaTransformation } from "@shared/media";

import type { RunPropertiesOptions } from "./properties";

/**
 * A tree node for SmartArt data.
 */
export interface SmartArtNode {
  text: string;
  children?: SmartArtNode[];
}

/**
 * Options for creating a SmartArtRun.
 *
 * @publicApi
 */
export interface SmartArtOptions {
  /** Tree-shaped data for the diagram content. */
  nodes: SmartArtNode[];
  /** Display dimensions */
  transformation: MediaTransformation;
  /** Floating positioning */
  floating?: Floating;
  /** Alternative text for accessibility */
  altText?: DocPropertiesOptions;
  /** wp:cNvGraphicFramePr locks (null = source had none; omit for default). */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Built-in layout ID ("default", "process1") or a custom layout definition. */
  layout?: string | LayoutDefinitionOptions;
  /** Built-in quick style ID ("simple1") or a custom style definition. */
  style?: string | StyleDefinitionOptions;
  /** Built-in color transform ID ("accent1_2") or a custom color definition. */
  color?: string | ColorDefinitionOptions;
  /** Run properties of the wrapping drawing run (round-trip fidelity). */
  runProperties?: RunPropertiesOptions;
  /** Word's pagination hint sharing the drawing run (round-trip fidelity). */
  lastRenderedPageBreak?: boolean;
  /**
   * Verbatim source-part bytes for byte-exact round-trip. When present,
   * generate re-emits these bytes under the diagrams part names instead of
   * rebuilding from nodes/layout/style/color (which stay populated and
   * readable); dropping raw rebuilds from the structured fields as usual.
   */
  raw?: SmartArtRawParts;
}
