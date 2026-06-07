import type { SlideChild } from "@file/slide/slide-child";

export interface GroupShapeOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly rotation?: number;
  readonly flipHorizontal?: boolean;
  readonly children: readonly SlideChild[];
}
