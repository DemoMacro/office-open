import type { UniversalMeasure } from "@office-open/core";
import type { FillOptions } from "@shared/drawingml/fill";
import type { OutlineOptions } from "@shared/drawingml/outline";

export interface LineShapeOptions {
  id?: number;
  name?: string;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  fill?: FillOptions;
  outline?: OutlineOptions;
}

export type ArrowheadType = "triangle" | "stealth" | "diamond" | "oval" | "open" | "none";

export interface ConnectorShapeOptions {
  id?: number;
  name?: string;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  fill?: FillOptions;
  outline?: OutlineOptions;
  beginArrowhead?: ArrowheadType;
  endArrowhead?: ArrowheadType;
  arrowheadWidth?: "small" | "medium" | "large";
  arrowheadLength?: "small" | "medium" | "large";
}
