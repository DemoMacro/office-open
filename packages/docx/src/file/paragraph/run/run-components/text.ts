/**
 * Text module for WordprocessingML run content.
 *
 * This module provides the buildText function for text content within runs.
 *
 * @module
 */
import { SpaceType } from "@file/shared";
import type { IXmlableObject } from "@file/xml-components";

interface TextOptions {
  space?: (typeof SpaceType)[keyof typeof SpaceType];
  text?: string;
}

/**
 * Builds a text element XML object.
 *
 * @param options - Text string or options with space/text
 * @returns IXmlableObject for the w:t element
 *
 * @internal
 */
export function buildText(options: string | TextOptions): IXmlableObject {
  if (typeof options === "string") {
    return {
      "w:t": [{ _attr: { "xml:space": SpaceType.PRESERVE } }, options],
    };
  }
  return {
    "w:t": [{ _attr: { "xml:space": options.space ?? SpaceType.DEFAULT } }, options.text ?? ""],
  };
}
