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
  relationships.elements?.find((e) => e.name === id)?.elements ?? [];

/**
 * Next sequential numeric id: the largest `attr` value among direct children
 * named `childName`, plus one. `seed` is the floor — e.g. PPTX `<p:sldId>`
 * values start at 255, so the first appended id is at least 256.
 *
 * Unifies the per-package `maxCommentId` / `maxSldId` / `maxSheetId` helpers.
 */
export const nextNumericId = (
  parent: Element | undefined,
  childName: string,
  attr: string,
  seed = 0,
): number => {
  let maxId = seed;
  for (const child of parent?.elements ?? []) {
    if (child.name !== childName) continue;
    const id = Number(child.attributes?.[attr]);
    if (Number.isFinite(id)) maxId = Math.max(maxId, id);
  }
  return maxId + 1;
};
