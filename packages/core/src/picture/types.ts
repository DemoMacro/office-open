import type { NonVisualDrawingPropertiesOptions } from "../drawing";
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
  /**
   * Image binary (base64 string, data URL, or Uint8Array). Absent on a
   * linked-only picture whose source is an external URL (no bytes in the
   * package) — that shape carries `sourceUrl` instead.
   */
  data?: DataType;
  /**
   * External image source URL (a:blip @r:link). A picture with a URL and no
   * {@link data} is linked-only — the owning part registers one External
   * image relationship instead of a media part. Both fields carry a local
   * cache plus its linked source.
   */
  sourceUrl?: string;
  /** Image format. Loose string here; each package narrows to its supported union. */
  type: string;
}
