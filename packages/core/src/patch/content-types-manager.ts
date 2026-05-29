/**
 * Content types manager for OOXML [Content_Types].xml management.
 */
import type { Element } from "@office-open/xml";

import { getFirstLevelElements } from "./xml-patch-utils";

export const appendContentType = (
  element: Element,
  contentType: string,
  extension: string,
): void => {
  const relationshipElements = getFirstLevelElements(element, "Types");

  const exist = relationshipElements.some(
    (el) =>
      el.type === "element" &&
      el.name === "Default" &&
      el?.attributes?.ContentType === contentType &&
      el?.attributes?.Extension === extension,
  );
  if (exist) {
    return;
  }

  relationshipElements.push({
    attributes: {
      ContentType: contentType,
      Extension: extension,
    },
    name: "Default",
    type: "element",
  });
};
