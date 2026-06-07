import type { FillOptions } from "@file/drawingml/fill";
import type { OutlineOptions } from "@file/drawingml/outline";

export interface LineShapeOptions {
  readonly id?: number;
  readonly name?: string;
  readonly x1?: number;
  readonly y1?: number;
  readonly x2?: number;
  readonly y2?: number;
  readonly fill?: FillOptions;
  readonly outline?: OutlineOptions;
}

export type ArrowheadType = "triangle" | "stealth" | "diamond" | "oval" | "open" | "none";

export interface ConnectorShapeOptions {
  readonly id?: number;
  readonly name?: string;
  readonly x1?: number;
  readonly y1?: number;
  readonly x2?: number;
  readonly y2?: number;
  readonly fill?: FillOptions;
  readonly outline?: OutlineOptions;
  readonly beginArrowhead?: ArrowheadType;
  readonly endArrowhead?: ArrowheadType;
  readonly arrowheadWidth?: "small" | "medium" | "large";
  readonly arrowheadLength?: "small" | "medium" | "large";
}
