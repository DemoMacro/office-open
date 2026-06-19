/**
 * Merge core-properties overrides onto an existing docProps/core.xml document.
 *
 * The patch layer unzips docProps/core.xml into an xml2js document structure
 * (its `.elements` holds the `cp:coreProperties` root). This helper parses the
 * existing values, layers `overrides` on top (override wins), and returns the
 * re-serialized XML — without the XML declaration, which the caller prepends.
 */
import type { Element } from "@office-open/xml";

import { buildCorePropertiesXmlString, parseCorePropsElement } from "../opc/core";
import type { CorePropertiesOptions } from "../opc/core";

export function applyCorePropertiesOverride(
  corePropsDoc: Element,
  overrides: Partial<CorePropertiesOptions>,
): string {
  const rootEl = (corePropsDoc.elements?.find((e) => e.name === "cp:coreProperties") ??
    corePropsDoc) as Element;
  const existing = parseCorePropsElement(rootEl);
  const merged: CorePropertiesOptions = { ...existing, ...overrides };
  return buildCorePropertiesXmlString(merged);
}
