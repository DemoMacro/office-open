/**
 * Symbol module for WordprocessingML run content.
 *
 * This module provides support for inserting symbol characters
 * from symbol fonts like Wingdings or Symbol.
 *
 * @module
 */

/**
 * Builds a symbol element XML object.
 *
 * @param char - Character code (hex value)
 * @param symbolfont - Symbol font name (default: "Wingdings")
 * @returns IXmlableObject for the w:sym element
 *
 * @internal
 */
export function buildSymbol(
  char: string = "",
  symbolfont: string = "Wingdings",
): Readonly<Record<string, unknown>> {
  return {
    "w:sym": { _attr: { "w:char": char, "w:font": symbolfont } },
  };
}
