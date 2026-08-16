/**
 * Black/white mode (a:ST_BlackWhiteMode) — the `@bwMode` container attribute on
 * shape properties (CT_ShapeProperties) and group shape properties
 * (CT_GroupShapeProperties). Controls how a drawing renders in black-and-white
 * view/print.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, ST_BlackWhiteMode
 *
 * @module
 */

/** ST_BlackWhiteMode — the 11 token values for the `@bwMode` attribute. */
export type BlackWhiteMode =
  | "clr"
  | "auto"
  | "gray"
  | "ltGray"
  | "invGray"
  | "grayWhite"
  | "blackGray"
  | "blackWhite"
  | "black"
  | "white"
  | "hidden";
