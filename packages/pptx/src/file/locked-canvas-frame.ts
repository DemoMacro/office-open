import type { TextBodyOptions } from "@file/shape/text-body";

export interface LockedCanvasShapeOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly geometry?: string;
  readonly fill?: string;
  readonly textBody?: TextBodyOptions;
}

export interface LockedCanvasFrameOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly id?: number;
  readonly name?: string;
  readonly children?: readonly LockedCanvasShapeOptions[];
}
