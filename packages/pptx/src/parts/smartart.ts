import type { UniversalMeasure } from "@office-open/core";
import type { GraphicFrameLockingOptions } from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions } from "@office-open/core/drawing";
import type {
  ColorDefinitionOptions,
  LayoutDefinitionOptions,
  StyleDefinitionOptions,
  TreeNode,
} from "@office-open/core/smartart";
import type { NvPrPlaceholderOptions } from "@parts/descriptors/graphic-frame";

export {
  getLayoutXml,
  getStyleXml,
  getColorXml,
  DEFAULT_DRAWING_XML,
  COLOR_CATEGORIES,
  LAYOUT_CATEGORIES,
  STYLE_CATEGORIES,
} from "@office-open/core/smartart";

export type { TreeNode };

/**
 * SmartArt diagram (p:graphicFrame). Extends NonVisualDrawingPropertiesOptions
 * (cNvPr authored like every drawing, survives cross-format copy);
 * `id`/`smartArtKey` auto-generated when omitted.
 *
 * @publicApi
 */
export interface SmartArtOptions extends NonVisualDrawingPropertiesOptions, NvPrPlaceholderOptions {
  /** Frame locking (a:graphicFrameLocks). undefined = fresh default
   * (noGrp="1"); null = empty cNvGraphicFramePr; object = explicit flags. */
  locking?: GraphicFrameLockingOptions | null;
  /** Diagram id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  /** Pre-generated SmartArt key (e.g. "smartart_1024"). Auto-generated if omitted. */
  smartArtKey?: string;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Tree-shaped data for the diagram content. */
  nodes: TreeNode[];
  /** Built-in layout ID ("default", "process1") or a custom layout definition. */
  layout?: string | LayoutDefinitionOptions;
  /** Built-in quick style ID ("simple1") or a custom style definition. */
  style?: string | StyleDefinitionOptions;
  /** Built-in color transform ID ("accent1_2") or a custom color definition. */
  color?: string | ColorDefinitionOptions;
}
