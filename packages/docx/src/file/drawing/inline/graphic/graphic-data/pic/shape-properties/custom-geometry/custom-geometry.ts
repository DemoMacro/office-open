/**
 * Custom geometry module for DrawingML shapes.
 *
 * Provides CT_CustomGeometry2D — user-defined 2D geometry with paths,
 * guides, adjust handles, connection sites, and a text insertion rectangle.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_CustomGeometry2D
 *
 * @module
 */

export { createCustomGeometry } from "@office-open/core/drawingml";
export type {
    CustomGeometryOptions,
    PathOptions,
    PathCommand,
    PathFillMode,
    ConnectionSite,
    GeomRect,
} from "@office-open/core/drawingml";
