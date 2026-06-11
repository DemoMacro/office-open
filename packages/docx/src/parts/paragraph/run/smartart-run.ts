/**
 * SmartArtRun types for WordprocessingML documents.
 *
 * SmartArt data is stored in `word/diagrams/data{n}.xml`.
 * Layout, style, and colors reference Word's built-in definitions.
 *
 * @module
 */
import type { Floating } from "@parts/drawing";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { MediaTransformation } from "@shared/media";

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
  /** Tree-shaped data for the diagram */
  data: {
    nodes: SmartArtNode[];
  };
  /** Display dimensions */
  transformation: MediaTransformation;
  /** Floating positioning */
  floating?: Floating;
  /** Alternative text for accessibility */
  altText?: DocPropertiesOptions;
  /** Layout ID (e.g. "default", "process1", "hierarchy1") */
  layout?: string;
  /** Quick style ID (e.g. "simple1", "moderate1") */
  style?: string;
  /** Color transform ID (e.g. "accent1_2", "colorful1") */
  color?: string;
}
