import type { FillOptions } from "@file/drawingml/fill";
import type { OutlineOptions } from "@file/drawingml/outline";

export interface LineShapeOptions {
  id?: number;
  name?: string;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  fill?: FillOptions;
  outline?: OutlineOptions;
}

export type ArrowheadType = "triangle" | "stealth" | "diamond" | "oval" | "open" | "none";

export interface ConnectorShapeOptions {
  id?: number;
  name?: string;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  fill?: FillOptions;
  outline?: OutlineOptions;
  beginArrowhead?: ArrowheadType;
  endArrowhead?: ArrowheadType;
  arrowheadWidth?: "small" | "medium" | "large";
  arrowheadLength?: "small" | "medium" | "large";
}
