/**
 * XML utility functions for patch operations.
 */
import { xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";

export const toJson = (xmlData: string): Element => {
  const xmlObj = xml2js(xmlData, {
    captureSpacesBetweenElements: true,
    compact: false,
  }) as Element;
  return xmlObj;
};

/**
 * Creates the inner content of a text element (`w:t` / `a:t`).
 *
 * Returns `[{ type: "text", text }]` for non-empty text, `[]` for empty.
 * The `xml:space` attribute is handled separately by `patchSpaceAttribute`.
 */
export const createTextElementContents = (text: string): Element[] =>
  text === "" ? [] : [{ text, type: "text" }];

export const patchSpaceAttribute = (element: Element): Element => ({
  ...element,
  attributes: {
    "xml:space": "preserve",
  },
});

export const getFirstLevelElements = (relationships: Element, id: string): Element[] =>
  relationships.elements?.filter((e) => e.name === id)[0].elements ?? [];
