import type { BaseGroupOptions, UniversalMeasure } from "@office-open/core";
import type { SlideChild } from "@parts/slide/slide-child";

/**
 * Group shape options for pptx slides (p:grpSp). The cNvPr fields
 * (name/description/title/hidden) come from {@link BaseGroupOptions}; the rest
 * is the pptx flat positioning model plus the group's children.
 */
export interface GroupOptions extends BaseGroupOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  rotation?: number;
  flipHorizontal?: boolean;
  children: SlideChild[];
}
