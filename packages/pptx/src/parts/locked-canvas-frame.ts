import type { UniversalMeasure } from "@office-open/core";
import type { TextBodyOptions } from "@shared/shape/text-body";

export interface LockedCanvasShapeOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  geometry?: string;
  fill?: string;
  textBody?: TextBodyOptions;
}

export interface LockedCanvasFrameOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  id?: number;
  name?: string;
  children?: LockedCanvasShapeOptions[];
}
