import type { PresetGeometryOptions, CustomGeometryOptions } from "@office-open/core/drawingml";

import type { EffectsOptions } from "./effects";
/**
 * Shape properties options type for PPTX.
 *
 * @module
 */
import type { FillOptions } from "./fill";
import type { OutlineOptions } from "./outline";

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
  effects?: EffectsOptions;
  connectionSites?: ConnectionSiteOptions[];
}
