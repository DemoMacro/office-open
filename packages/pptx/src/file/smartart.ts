import type { TreeNode } from "@office-open/core/smartart";

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

export interface SmartArtOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  nodes: TreeNode[];
  name?: string;
  /** Layout ID (e.g. "default", "process1", "hierarchy1") */
  layout?: string;
  /** Quick style ID (e.g. "simple1", "moderate1") */
  style?: string;
  /** Color transform ID (e.g. "accent1_2", "colorful1") */
  color?: string;
}
