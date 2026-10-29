/**
 * SmartArtRun types for WordprocessingML documents.
 *
 * SmartArt data is stored in `word/diagrams/data{n}.xml`.
 * Layout, style, and colors reference Word's built-in definitions or carry
 * full custom definitions.
 *
 * @module
 */
import type { LongHexNumber } from "@office-open/core";
import type { TreeNode } from "@office-open/core/smartart";
import type {
  ColorDefinitionOptions,
  DataModelOptions,
  LayoutDefinitionOptions,
  SmartArtRawParts,
  StyleDefinitionOptions,
} from "@office-open/core/smartart";
import type { Floating } from "@parts/drawing";
import type { DrawingExtensionIds, GraphicFrameLocksOptions } from "@parts/drawing/descriptor";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { MediaTransformation } from "@shared/media";

import type { RunPropertiesOptions } from "./properties";

export type { TreeNode };

/**
 * Options for creating a SmartArtRun.
 *
 * @publicApi
 */
export interface SmartArtOptions {
  /** Tree-shaped data for the diagram content. */
  nodes: TreeNode[];
  /** Complete data model for points, connections, properties, and metadata. */
  dataModel?: DataModelOptions;
  /** Display dimensions */
  transformation: MediaTransformation;
  /** Floating positioning */
  floating?: Floating;
  /** Alternative text for accessibility */
  altText?: DocPropertiesOptions;
  /** wp:cNvGraphicFramePr locks (null = source had none; omit for default). */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Word 2010 wrapper ids; round-trip only. */
  extensionIds?: DrawingExtensionIds;
  /** Built-in layout ID ("default", "process1") or a custom layout definition. */
  layout?: string | LayoutDefinitionOptions;
  /** Built-in quick style ID ("simple1") or a custom style definition. */
  style?: string | StyleDefinitionOptions;
  /** Built-in color transform ID ("accent1_2") or a custom color definition. */
  color?: string | ColorDefinitionOptions;
  /** Run properties of the wrapping drawing run (round-trip fidelity). */
  runProperties?: RunPropertiesOptions;
  /** Revision save ID of the wrapping w:r (w:rsidR, round-trip). */
  additionRsid?: LongHexNumber;
  /** Revision save ID when the wrapping w:r was deleted (w:rsidDel, round-trip). */
  deletionRsid?: LongHexNumber;
  /** Revision save ID of the wrapping w:r properties (w:rsidRPr, round-trip). */
  runPropertiesRsid?: LongHexNumber;
  /** Word's pagination hint sharing the drawing run (round-trip fidelity). */
  lastRenderedPageBreak?: boolean;
  /**
   * Verbatim source-part bytes for byte-exact round-trip: when present,
   * generate re-emits them instead of rebuilding (nodes/layout/style/color
   * stay populated and readable).
   */
  raw?: SmartArtRawParts;
}
