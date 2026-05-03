/**
 * Outline (line) properties for DrawingML shapes.
 *
 * This module provides support for configuring outline properties including
 * width, cap style, compound line types, fill properties, dash, and join.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_LineProperties
 *
 * @module
 */

export { createOutline } from "@office-open/core/drawingml";
export type { OutlineOptions, OutlineFillProperties } from "@office-open/core/drawingml";
export {
    LineCap,
    CompoundLine,
    PenAlignment,
    PresetDash,
    LineJoin,
} from "@office-open/core/drawingml";
