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
  StyleDefinitionOptions,
} from "@office-open/core/smartart";
import type { Floating } from "@parts/drawing";
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
}
