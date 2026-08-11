import type {
  EndpointConnectionOptions,
  ConnectorLockingOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { OutlineOptions } from "@office-open/core/drawingml";
import type { FillOptions } from "@shared/drawingml/fill";

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
  /** a:cxnSpLocks — connector locking (inside p:cNvCxnSpPr). */
  locking?: ConnectorLockingOptions;
  /** a:stCxn — start endpoint glued to a shape connection site. */
  startConnection?: EndpointConnectionOptions;
  /** a:endCxn — end endpoint glued to a shape connection site. */
  endConnection?: EndpointConnectionOptions;
}
