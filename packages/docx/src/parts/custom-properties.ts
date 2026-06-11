/**
 * Custom Properties module for WordprocessingML documents.
 *
 * Provides the CustomPropertyOptions interface.
 * XML generation is handled by the descriptor pipeline (customPropertiesDesc).
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCustom.xsd
 *
 * @module
 */

/**
 * Options for creating a custom property.
 *
 * @property name - The property name
 * @property value - The property value (as string)
 */
export interface CustomPropertyOptions {
  /** The property name */
  name: string;
  /** The property value (as string) */
  value: string;
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, textOf } from "@office-open/xml";

export interface CustomPropertiesInput {
  properties: CustomPropertyOptions[];
}

export const customPropertiesDesc: CustomDescriptor<CustomPropertiesInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"' +
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    ];
    let pid = 2;
    for (const prop of opts.properties) {
      p.push(
        `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${escapeXml(prop.name)}">` +
          `<vt:lpwstr>${escapeXml(prop.value)}</vt:lpwstr>` +
          `</property>`,
      );
      pid++;
    }
    p.push("</Properties>");
    return p.join("");
  },

  parse(el, _ctx) {
    const properties: CustomPropertyOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "property") continue;
      const name = attr(child, "name");
      if (!name) continue;

      // Find the value element (vt:lpwstr, vt:lpstr, vt:i4, etc.)
      const valueEl = child.elements?.find((e) => e.name?.startsWith("vt:"));
      const value = valueEl ? (textOf(valueEl) ?? "") : "";
      properties.push({ name, value });
    }
    return { properties } as Record<string, unknown>;
  },
};
