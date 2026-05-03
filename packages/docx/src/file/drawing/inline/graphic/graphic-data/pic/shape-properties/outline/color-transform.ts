/**
 * Color transform elements for DrawingML colors.
 *
 * This module provides color transformation elements defined in EG_ColorTransform,
 * which can be applied as child elements to any color type (srgbClr, schemeClr, etc.).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, EG_ColorTransform
 *
 * @module
 */

export { createColorTransforms } from "@office-open/core/drawingml";
export type { ColorTransformOptions } from "@office-open/core/drawingml";
