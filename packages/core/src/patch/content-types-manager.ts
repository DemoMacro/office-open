/**
 * Content types manager for OOXML [Content_Types].xml management.
 */
import type { Element } from "@office-open/xml";

import { PatchError } from "./diagnostics";

/**
 * The `<Types>` child array. Returns `undefined` when `element` has no
 * `<Types>` child (malformed `[Content_Types].xml`); public operations reject
 * that state with an `invalid-part` diagnostic.
 */
const typesChildren = (element: Element): Element[] | undefined => {
  const types = element.elements?.find((e) => e.name === "Types");
  if (!types) return undefined;
  return types.elements ?? (types.elements = []);
};

const typesOrThrow = (element: Element, operation: string): Element[] => {
  const types = typesChildren(element);
  if (!types) {
    throw new PatchError({
      message: "Malformed [Content_Types].xml: missing Types root",
      part: "[Content_Types].xml",
      path: "/Types",
      operation,
      reason: "invalid-part",
    });
  }
  return types;
};

export const appendContentType = (
  element: Element,
  contentType: string,
  extension: string,
): void => {
  const els = typesOrThrow(element, "append-default");

  // OPC Default Extension matching is case-insensitive (ECMA-376-2 §10.1.2):
  // an existing `JPG` must block a colliding `jpg`, or Word rejects the package.
  const exist = els.some(
    (el) =>
      el.type === "element" &&
      el.name === "Default" &&
      String(el?.attributes?.Extension ?? "").toLowerCase() === extension.toLowerCase(),
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
  const els = typesOrThrow(element, "append-override");

  // OPC Override PartName matching is case-insensitive (ECMA-376-2 §10.1.4).
  const exists = els.some(
    (el) =>
      el.type === "element" &&
      el.name === "Override" &&
      String(el?.attributes?.PartName ?? "").toLowerCase() === partName.toLowerCase(),
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

/**
 * Remove an `<Override>` from a `[Content_Types].xml` root by `PartName`
 * (case-insensitive). Sibling to {@link appendOverride} for patch paths that
 * drop a part from the package.
 */
export const removeOverride = (element: Element, partName: string): void => {
  const els = typesOrThrow(element, "remove-override");

  const kept = els.filter(
    (el) =>
      !(
        el.type === "element" &&
        el.name === "Override" &&
        String(el?.attributes?.PartName ?? "").toLowerCase() === partName.toLowerCase()
      ),
  );
  if (kept.length !== els.length) {
    element.elements!.find((e) => e.name === "Types")!.elements = kept;
  }
};
