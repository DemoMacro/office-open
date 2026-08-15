/**
 * Preset geometry stringifier for DrawingML shapes.
 *
 * @module
 */

import type { GeometryGuide } from "./adjustment-values";
import { stringifyAdjustmentValues } from "./adjustment-values";

export interface PresetGeometryOptions {
  preset?: string;
  adjustmentValues?: readonly GeometryGuide[];
}

/**
 * Build a:prstGeom XML string.
 */
export function stringifyPresetGeometry(options?: PresetGeometryOptions): string {
  const prst = options?.preset ?? "rect";
  const avLst = stringifyAdjustmentValues(options?.adjustmentValues);
  return `<a:prstGeom prst="${prst}">${avLst}</a:prstGeom>`;
}
