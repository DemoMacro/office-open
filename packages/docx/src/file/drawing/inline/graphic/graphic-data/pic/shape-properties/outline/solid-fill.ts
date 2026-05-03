/**
 * Solid fill element for DrawingML shapes.
 *
 * This module provides solid fill support for outlines and shapes,
 * supporting RGB, scheme, HSL, system, and preset colors.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_SolidColorFillProperties
 *
 * @module
 */

export { createColorElement, createSolidFill } from "@office-open/core/drawingml";
export type { SolidFillOptions } from "@office-open/core/drawingml";
