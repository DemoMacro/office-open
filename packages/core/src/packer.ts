/**
 * Shared packer utilities for OOXML document generation.
 *
 * @module
 */

export interface XmlifyedFile {
  readonly path: string;
  readonly data: string | Uint8Array;
}

export const PrettifyType = {
  NONE: "",
  WITH_2_BLANKS: "  ",
  WITH_4_BLANKS: "    ",
  WITH_TAB: "\t",
} as const;

export const convertPrettifyType = (
  prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
): (typeof PrettifyType)[keyof typeof PrettifyType] | undefined =>
  prettify === true ? PrettifyType.WITH_2_BLANKS : prettify === false ? undefined : prettify;
