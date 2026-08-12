import type {
  BaseConnectorOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { OutlineOptions } from "@office-open/core/drawingml";
import type { FillOptions } from "@shared/drawingml/fill";

export interface LineShapeOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  fill?: FillOptions;
  outline?: OutlineOptions;
}

export interface ConnectorShapeOptions extends BaseConnectorOptions {
  id?: number;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  fill?: FillOptions;
  outline?: OutlineOptions;
}
