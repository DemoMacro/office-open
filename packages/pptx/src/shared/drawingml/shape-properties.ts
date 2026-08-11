import type {
  PresetGeometryOptions,
  CustomGeometryOptions,
  OutlineOptions,
  EffectListOptions,
} from "@office-open/core/drawingml";

/**
 * Shape properties options type for PPTX.
 *
 * @module
 */
import type { FillOptions } from "./fill";

export interface ConnectionSiteOptions {
  x: number;
  y: number;
  angle?: number;
}

export interface ShapePropertiesOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  flipHorizontal?: boolean;
  rotation?: number;
  geometry?: string | PresetGeometryOptions;
  customGeometry?: CustomGeometryOptions;
  fill?: FillOptions;
  outline?: OutlineOptions;
  effects?: EffectListOptions;
  connectionSites?: ConnectionSiteOptions[];
}
