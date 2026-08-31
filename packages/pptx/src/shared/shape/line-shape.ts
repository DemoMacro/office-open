import type {
  BaseConnectorOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";
import type {
  PresetGeometryOptions,
  ShapeLockingOptions,
  ShapePropertiesOptions,
  ShapeType,
  TextBodyOptions,
} from "@office-open/core/drawing";

import type { ShapeStyleOptions } from "./shape";

/** spPr paint children carried by endpoint-model shapes (p:sp line, p:cxnSp). */
export type EndpointShapeProperties = Pick<
  ShapePropertiesOptions,
  "fill" | "outline" | "effects" | "scene3d" | "shape3d"
>;

/** Connector paint adds the preset geometry to the shared endpoint paint. */
export type ConnectorShapeProperties = EndpointShapeProperties & {
  /** Connector preset geometry (a:prstGeom @prst with adjustment guides). */
  geometry?: ShapeType | PresetGeometryOptions;
};

export interface LineShapeOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
  /** Shape locks (a:spLocks inside p:cNvSpPr). */
  locking?: ShapeLockingOptions;
  /** Text body — source lines carry wrap/anchor hints and an empty paragraph. */
  textBody?: TextBodyOptions;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  /**
   * Line paint (a:spPr children): fill/outline/effects/3D. Endpoints stay
   * top-level — direction is encoded as xfrm flip, not an owner transform.
   */
  properties?: EndpointShapeProperties;
  /** Shape style matrix reference (p:style). */
  style?: ShapeStyleOptions;
}

/**
 * Connector options for pptx slides (p:cxnSp). The cNvPr + locking + endpoint
 * connection fields come from `BaseConnectorOptions`; the rest is the pptx
 * two-endpoint positioning model plus the connector paint.
 */
export interface ConnectorOptions extends BaseConnectorOptions {
  id?: number;
  /**
   * Connector paint (a:spPr children): geometry/fill/outline/effects/3D. The
   * endpoint model defaults to "line"; source connectors often use
   * bentConnector/elbowConnector forms with adjusted values. Endpoints stay
   * top-level.
   */
  properties?: ConnectorShapeProperties;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  /** Shape style matrix reference (p:style). */
  style?: ShapeStyleOptions;
}
