import type { DataType, UniversalMeasure } from "@office-open/core";

export interface PictureOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  data: DataType;
  type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
  name?: string;
}
