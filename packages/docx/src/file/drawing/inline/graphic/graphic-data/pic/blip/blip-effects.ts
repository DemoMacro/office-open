/**
 * Blip image adjustment effects for DrawingML.
 *
 * These effects are applied directly to the image data within the `<a:blip>` element,
 * corresponding to Word's "Picture Format > Adjust" features (brightness, contrast,
 * grayscale, tint, duotone, etc.).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Blip children
 *
 * @module
 */

export { createBlipEffects } from "@office-open/core/drawingml";
export type { BlipEffectsOptions } from "@office-open/core/drawingml";
