/**
 * Text element descriptors — re-exported from core DrawingML.
 *
 * PPTX text runs/paragraphs/fields (a:r/a:p/a:fld) share one descriptor
 * implementation across formats via core. This module re-exports it so internal
 * callers (shape.ts, table.ts) keep their import paths.
 *
 * @module
 */

export { runPropertiesDesc, textRunDesc, paragraphDesc } from "@office-open/core/drawing";

export type {
  ParagraphDescriptorOptions,
  TextBreakOptions,
  BulletAutoNumOptions,
  BulletCharOptions,
  BulletOptions,
  TextParagraphPropertiesOptions,
  TextRunOptions,
  TextHyperlinkOptions,
  TextCharacterPropertiesOptions,
} from "@office-open/core/drawing";
