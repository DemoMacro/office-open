/**
 * Custom Properties module — shared OPC part (docProps/custom.xml).
 *
 * Format-agnostic: CT_CustomProperties (ISO-IEC29500-4_2016 shared-documentPropertiesCustom.xsd)
 * is identical across docx/pptx/xlsx.
 *
 * @module
 */

import { attr, escapeXml } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import { parseVariantValue, stringifyVariantValue, type VariantValue } from "./variant-types";

/**
 * Options for a single custom property.
 *
 * @property name - The property name
 * @property value - The property value; the JS type picks the vt:* element
 * (string → lpwstr, number → i4/r8, boolean → bool, Date → filetime)
 */
export interface CustomPropertyOptions {
  /** The property name */
  name: string;
  /** The property value (as string) */
  value: VariantValue;
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
          stringifyVariantValue(prop.value) +
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

      // The value element is a single vt:* scalar (lpwstr, i4, bool, …)
      const valueEl = child.elements?.find((e) => e.name?.startsWith("vt:"));
      if (valueEl) {
        const value = parseVariantValue(valueEl);
        if (value !== undefined) properties.push({ name, value });
      }
    }
    return { properties };
  },
};
