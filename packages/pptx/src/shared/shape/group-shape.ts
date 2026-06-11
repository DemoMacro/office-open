import type { UniversalMeasure } from "@office-open/core";
import type { SlideChild } from "@parts/slide/slide-child";

export interface GroupShapeOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  rotation?: number;
  flipHorizontal?: boolean;
  children: SlideChild[];
}
