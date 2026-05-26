import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { textOf } from "@office-open/xml";

/**
 * Parse docProps/custom.xml into custom property entries.
 *
 * XML format:
 * ```xml
 * <property name="..." fmtid="..." pid="2">
 *   <vt:lpwstr>value</vt:lpwstr>
 * </property>
 * ```
 */
export function parseCustomProperties(el: Element | undefined): { name: string; value: string }[] {
  if (!el) return [];

  const result: { name: string; value: string }[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "property") continue;
    const name = attr(child, "name");
    if (!name) continue;

    // Find the value element (vt:lpwstr, vt:lpstr, etc.)
    const valueEl = child.elements?.find((e) => e.name?.startsWith("vt:"));
    const value = valueEl ? (textOf(valueEl) ?? "") : "";
    result.push({ name, value });
  }

  return result;
}
