import type { SlideChild } from "@file/slide/slide-child";

export interface GroupShapeOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
  flipHorizontal?: boolean;
  children: SlideChild[];
}
