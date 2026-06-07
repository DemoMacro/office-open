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
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly nodes: readonly TreeNode[];
  readonly name?: string;
  /** Layout ID (e.g. "default", "process1", "hierarchy1") */
  readonly layout?: string;
  /** Quick style ID (e.g. "simple1", "moderate1") */
  readonly style?: string;
  /** Color transform ID (e.g. "accent1_2", "colorful1") */
  readonly color?: string;
}
