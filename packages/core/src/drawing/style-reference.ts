/**
 * Theme style-matrix reference for DrawingML.
 *
 * Reference: ISO/IEC 29500-1, dml-main.xsd, CT_StyleMatrixReference
 *
 * @module
 */
import type { SolidFillOptions } from "./color/solid-fill";

/**
 * Reference into the theme style matrix — a:lnRef / a:fillRef / a:effectRef.
 *
 * Shared by theme object defaults (spDef style), shape styles, and table
 * styles; declared once here so every consumer sees the same shape.
 */
export interface StyleMatrixReferenceOptions {
  /** Index into the style matrix list (idx attribute). */
  index: number;
  /** Color component (EG_ColorChoice). */
  color?: SolidFillOptions;
}
