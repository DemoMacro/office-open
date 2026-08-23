import type {
  BaseConnectorOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";
import type {
  EffectListOptions,
  OutlineOptions,
  PresetGeometryOptions,
  Scene3DOptions,
  Shape3DOptions,
  ShapeLockingOptions,
  ShapeType,
  TextBodyOptions,
} from "@office-open/core/drawing";
import type { FillOptions } from "@shared/drawing/fill";

import type { ShapeStyleOptions } from "./shape";

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
  fill?: FillOptions;
  outline?: OutlineOptions;
  /** Effect list (a:effectLst) inside spPr. An empty object emits the bare element. */
  effects?: EffectListOptions;
  /** 3D scene (a:scene3d) inside spPr. */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (a:sp3d) inside spPr. */
  shape3d?: Shape3DOptions;
  /** Shape style matrix reference (p:style). */
  style?: ShapeStyleOptions;
}

/**
 * Connector options for pptx slides (p:cxnSp). The cNvPr + locking + endpoint
 * connection fields come from `BaseConnectorOptions`; the rest is the
 * pptx two-endpoint positioning model plus line fill/outline.
 */
export interface ConnectorOptions extends BaseConnectorOptions {
  id?: number;
  /**
   * Connector preset geometry (a:prstGeom @prst with optional adjustment
   * guides). The endpoint model defaults to "line"; source connectors often
   * use bentConnector/elbowConnector forms with adjusted values.
   */
  geometry?: ShapeType | PresetGeometryOptions;
  x1?: number | UniversalMeasure;
  y1?: number | UniversalMeasure;
  x2?: number | UniversalMeasure;
  y2?: number | UniversalMeasure;
  fill?: FillOptions;
  outline?: OutlineOptions;
  /** Effect list (a:effectLst) inside spPr. An empty object emits the bare element. */
  effects?: EffectListOptions;
  /** 3D scene (a:scene3d) inside spPr. */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (a:sp3d) inside spPr. */
  shape3d?: Shape3DOptions;
  /** Shape style matrix reference (p:style). */
  style?: ShapeStyleOptions;
}
