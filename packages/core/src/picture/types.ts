import type { NonVisualDrawingPropertiesOptions } from "../drawingml";
import type { DataType } from "../util/data-type";

/**
 * Base picture options — the shared shape across docx/pptx/xlsx pictures.
 *
 * Carries the binary payload (data/type) plus the non-visual drawing
 * properties (name/description/title/hidden) that mirror a:CT_NonVisualDrawingProps,
 * the cNvPr/docPr type every package's picture emits. Positioning is
 * package-specific (docx transformation, pptx x/y, xlsx cell anchor) and lives
 * on each package's PictureOptions, not here.
 *
 * docx does not extend this base (its PictureOptions is a discriminated union
 * by image format); the cross-format converter reads the equivalent fields from
 * docx altText instead.
 */
export interface BasePictureOptions extends NonVisualDrawingPropertiesOptions {
  /** Image binary (base64 string, data URL, or Uint8Array). */
  data: DataType;
  /** Image format. Loose string here; each package narrows to its supported union. */
  type: string;
}
