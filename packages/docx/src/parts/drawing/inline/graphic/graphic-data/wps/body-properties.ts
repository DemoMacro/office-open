/**
 * Text body properties — re-exported from core DrawingML.
 *
 * DOCX wraps bodyPr in wps:bodyPr (WordprocessingML drawing namespace); this
 * module re-exports core's CT_TextBodyProperties implementation and fixes the
 * tag at wps:bodyPr for DOCX callers.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TextBodyProperties
 *
 * @module
 */

export {
  VerticalAnchor,
  TextVertOverflowType,
  TextHorzOverflowType,
  TextVerticalType,
  TextBodyWrappingType,
  parseBodyProperties,
  bodyPropertiesDesc,
} from "@office-open/core/drawingml";
export type {
  NormalAutofitOptions,
  PresetTextShapeOptions,
  FlatTextOptions,
  BodyPropertiesOptions,
} from "@office-open/core/drawingml";

import type { BodyPropertiesOptions } from "@office-open/core/drawingml";
import { createBodyProperties as coreCreateBodyProperties } from "@office-open/core/drawingml";

/**
 * Create a wps:bodyPr element. DOCX uses the wps: prefix (WordprocessingML
 * drawing); core's default is a:bodyPr (PPTX/XLSX).
 */
export const createBodyProperties = (options: BodyPropertiesOptions = {}): string =>
  coreCreateBodyProperties(options, "wps:bodyPr");
