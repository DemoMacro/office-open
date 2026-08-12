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

/**
 * Connector options for pptx slides (p:cxnSp). The cNvPr + locking + endpoint
 * connection fields come from {@link BaseConnectorOptions}; the rest is the
 * pptx two-endpoint positioning model plus line fill/outline.
 */
export interface ConnectorOptions extends BaseConnectorOptions {
  id?: number;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  fill?: FillOptions;
  outline?: OutlineOptions;
}

/**
 * @deprecated Use {@link ConnectorOptions}. Kept as an alias for backward
 * compatibility; matches the OOXML element name (p:cxnSp → Connector) across packages.
 */
export type ConnectorShapeOptions = ConnectorOptions;
