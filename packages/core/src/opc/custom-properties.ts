/**
 * Custom Properties module — shared OPC part (docProps/custom.xml).
 *
 * Format-agnostic: CT_CustomProperties (ISO-IEC29500-4_2016 shared-documentPropertiesCustom.xsd)
 * is identical across docx/pptx/xlsx.
 *
 * @module
 */

import { attr, escapeXml, textOf } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";

/**
 * Options for a single custom property.
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

/** Input shape for the custom-properties descriptor. */
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
    return { properties };
  },
};
