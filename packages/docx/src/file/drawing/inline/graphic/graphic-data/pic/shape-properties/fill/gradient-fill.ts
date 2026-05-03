/**
 * Gradient fill element for DrawingML shapes.
 *
 * This module provides gradient fill support with linear and path shading.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GradientFillProperties
 *
 * @module
 */

export { createGradientFill, createGradientStop } from "@office-open/core/drawingml";
export type {
    GradientFillOptions,
    GradientShadeOptions,
    LinearShadeOptions,
    PathShadeOptions,
    IGradientStop,
    RelativeRect,
} from "@office-open/core/drawingml";
export { PathShadeType, TileFlipMode } from "@office-open/core/drawingml";
