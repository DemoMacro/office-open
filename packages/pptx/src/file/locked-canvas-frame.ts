import type { TextBodyOptions } from "@file/shape/text-body";

export interface LockedCanvasShapeOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  geometry?: string;
  fill?: string;
  textBody?: TextBodyOptions;
}

export interface LockedCanvasFrameOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  id?: number;
  name?: string;
  children?: LockedCanvasShapeOptions[];
}
