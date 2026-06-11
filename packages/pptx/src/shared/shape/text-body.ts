/**
 * Text body options type for PPTX shapes.
 *
 * @module
 */
import type { UniversalMeasure } from "@office-open/core";

import type { VerticalAlignment } from "../table/table-cell";
import type { ParagraphOptions } from "./paragraph/paragraph";

export interface TextBodyOptions {
  text?: string;
  children?: (ParagraphOptions | string)[];
  vertical?:
    | "horz"
    | "vert"
    | "vert270"
    | "wordArtVert"
    | "eaVert"
    | "mongolianVert"
    | "wordArtVertRtl";
  anchor?: VerticalAlignment;
  autoFit?: "normal" | "shape" | "none";
  wrap?: "square" | "none";
  margins?: {
    top?: number | UniversalMeasure;
    bottom?: number | UniversalMeasure;
    left?: number | UniversalMeasure;
    right?: number | UniversalMeasure;
  };
  marginTop?: number | UniversalMeasure;
  marginBottom?: number | UniversalMeasure;
  columns?: number;
  columnSpacing?: number;
}
