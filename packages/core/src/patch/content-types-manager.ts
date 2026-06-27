/**
 * Content types manager for OOXML [Content_Types].xml management.
 */
import type { Element } from "@office-open/xml";

/**
 * The `<Types>` child array, initialized in place when absent (so a `<Default>`
 * or `<Override>` can always be appended). Returns `undefined` if `element` has
 * no `<Types>` child (malformed `[Content_Types].xml`).
 */
const typesChildren = (element: Element): Element[] | undefined => {
  const types = element.elements?.find((e) => e.name === "Types");
  if (!types) return undefined;
  return types.elements ?? (types.elements = []);
};

export const appendContentType = (
  element: Element,
  contentType: string,
  extension: string,
): void => {
  const els = typesChildren(element);
  if (!els) return;

  const exist = els.some(
    (el) =>
      el.type === "element" &&
      el.name === "Default" &&
      el?.attributes?.ContentType === contentType &&
      el?.attributes?.Extension === extension,
  );
  if (exist) {
    return;
  }

  els.push({
    attributes: {
      ContentType: contentType,
      Extension: extension,
    },
    name: "Default",
    type: "element",
  });
};

/**
 * Append an `<Override>` element to a `[Content_Types].xml` root, deduped by
 * `PartName`. Sibling to {@link appendContentType} (which handles `<Default>`):
 * unifies the per-package Override-dedup blocks in docx/pptx/xlsx patchers.
 */
export const appendOverride = (element: Element, partName: string, contentType: string): void => {
  const els = typesChildren(element);
  if (!els) return;

  const exists = els.some(
    (el) =>
      el.type === "element" && el.name === "Override" && el?.attributes?.PartName === partName,
  );
  if (exists) return;

  els.push({
    attributes: {
      PartName: partName,
      ContentType: contentType,
    },
    name: "Override",
    type: "element",
  });
};
