/**
 * Built-in SmartArt layout, style, and color URIs.
 *
 * These reference Word's built-in definitions — we only need to point to them.
 *
 * @module
 */

/** Built-in SmartArt layout URIs */
export const LAYOUTS: Record<string, string> = {
    process: "http://schemas.openxmlformats.org/drawingml/2006/diagram/process1",
    hierarchy: "http://schemas.openxmlformats.org/drawingml/2006/diagram/hierarchy1",
    cycle: "http://schemas.openxmlformats.org/drawingml/2006/diagram/cycle1",
    pyramid: "http://schemas.openxmlformats.org/drawingml/2006/diagram/pyramid1",
    list: "http://schemas.openxmlformats.org/drawingml/2006/diagram/list1",
};

/** Default style URI (accent 1) */
export const DEFAULT_STYLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagramstyle/2";

/** Default color transform URI (colorful accent 1) */
export const DEFAULT_COLOR_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagramcolor/3";
