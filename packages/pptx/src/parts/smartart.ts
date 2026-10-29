import type { Transform2DOptions } from "@office-open/core";
import type { GraphicFrameLockingOptions } from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions } from "@office-open/core/drawing";
import type {
  ColorDefinitionOptions,
  DataModelOptions,
  LayoutDefinitionOptions,
  StyleDefinitionOptions,
  SmartArtRawParts,
  TreeNode,
} from "@office-open/core/smartart";
import type { GraphicFrameNvPrOptions } from "@parts/descriptors/graphic-frame";

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

/** One source relationship from `ppt/diagrams/_rels/dataN.xml.rels`. */
export interface SmartArtDataRelationshipOptions {
  /** Source relationship Id (`rId3`). */
  id: string;
  /** Full OPC relationship Type URI. */
  relationshipType: string;
  /** Relationship Target, including relative paths and URLs. */
  target: string;
  /** TargetMode for external URLs. */
  targetMode?: "External";
}

/**
 * SmartArt diagram (p:graphicFrame). Extends NonVisualDrawingPropertiesOptions
 * (cNvPr authored like every drawing, survives cross-format copy);
 * `id`/`smartArtKey` auto-generated when omitted.
 *
 * @publicApi
 */
export interface SmartArtOptions
  extends NonVisualDrawingPropertiesOptions, GraphicFrameNvPrOptions, Transform2DOptions {
  /** Frame locking (a:graphicFrameLocks). undefined = fresh default
   * (noGrp="1"); null = empty cNvGraphicFramePr; object = explicit flags. */
  locking?: GraphicFrameLockingOptions | null;
  /** Diagram id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  /** Pre-generated SmartArt key (e.g. "smartart_1024"). Auto-generated if omitted. */
  smartArtKey?: string;
  /** Tree-shaped data for the diagram content. */
  nodes: TreeNode[];
  /** Complete canonical data model; preferred when rebuilding without raw bytes. */
  dataModel?: DataModelOptions;
  /** Built-in layout ID ("default", "process1") or a custom layout definition. */
  layout?: string | LayoutDefinitionOptions;
  /** Built-in quick style ID ("simple1") or a custom style definition. */
  style?: string | StyleDefinitionOptions;
  /** Built-in color transform ID ("accent1_2") or a custom color definition. */
  color?: string | ColorDefinitionOptions;
  /** Verbatim source diagram parts. Round-trip only — do not hand-author. */
  raw?: SmartArtRawParts;
  /** Source data-part relationships. Round-trip only — do not hand-author. */
  sourceDataRelationships?: SmartArtDataRelationshipOptions[];
  /** Source dgm:relIds relationship Ids (r:dm/r:lo/r:qs/r:cs). Round-trip only. */
  sourceRelIds?: { dm?: string; lo?: string; qs?: string; cs?: string };
  /** Source diagram part paths relative to the package root. Round-trip only. */
  sourcePaths?: { data?: string; layout?: string; quickStyle?: string; colors?: string };
}
