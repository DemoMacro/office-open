/**
 * Custom dash pattern for DrawingML outlines.
 *
 * This module provides support for custom dash patterns defined by
 * a list of dash stops, each specifying dash and space lengths as
 * percentages relative to line width.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_DashStopList, CT_DashStop
 *
 * @module
 */

export { createCustomDash } from "@office-open/core/drawingml";
export type { DashStop } from "@office-open/core/drawingml";
