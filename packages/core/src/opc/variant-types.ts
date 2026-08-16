/**
 * Variant value types (vt:*) — shared vocabulary for the docProps parts.
 *
 * custom.xml property values and app.xml vectors (HeadingPairs,
 * TitlesOfParts) are expressed with the vt:* element set defined in
 * shared-documentPropertiesVariantTypes.xsd. JS values map as:
 * string → vt:lpwstr, boolean → vt:bool, Date → vt:filetime, integers →
 * vt:i4, other numbers → vt:r8.
 *
 * The two domains spell strings differently and Excel enforces the split:
 * app.xml vectors (CT_VectorLpstr) must carry vt:lpstr — Excel refuses to
 * open a file whose HeadingPairs/TitlesOfParts use vt:lpwstr — while
 * custom.xml property values use vt:lpwstr, the spelling Office itself
 * writes there.
 *
 * @module
 */

import { escapeXml, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { parseOnOff } from "../util/values";

/** JS-level variant value carried by a vt:* element. */
export type VariantValue = string | number | boolean | Date;

/** Numeric vt:* tags parsed back to a JS number. */
const NUMERIC_VT_TAGS = new Set([
  "vt:i1",
  "vt:i2",
  "vt:i4",
  "vt:i8",
  "vt:ui1",
  "vt:ui2",
  "vt:ui4",
  "vt:ui8",
  "vt:int",
  "vt:uint",
  "vt:r4",
  "vt:r8",
  "vt:decimal",
]);

/**
 * Serialize a JS value as its vt:* element (stringify side, custom.xml domain).
 *
 * Strings stay lpwstr (the spelling Office writes for custom property
 * values), integers become i4, other numbers r8, booleans bool, and Dates
 * filetime (xsd:dateTime form).
 */
export function stringifyVariantValue(value: VariantValue): string {
  if (typeof value === "string") return `<vt:lpwstr>${escapeXml(value)}</vt:lpwstr>`;
  if (typeof value === "boolean") return `<vt:bool>${value ? "true" : "false"}</vt:bool>`;
  if (value instanceof Date) return `<vt:filetime>${value.toISOString()}</vt:filetime>`;
  return Number.isInteger(value) ? `<vt:i4>${value}</vt:i4>` : `<vt:r8>${value}</vt:r8>`;
}

/**
 * Parse a vt:* scalar element back to a JS value (parse side).
 *
 * Lenient about unknown tags: they round-trip as their text content, the
 * same fallback the previous string-only custom-properties parse used.
 */
export function parseVariantValue(el: Element): VariantValue | undefined {
  const text = textOf(el);
  switch (el.name) {
    case "vt:lpwstr":
    case "vt:lpstr":
    case "vt:bstr":
      return text;
    case "vt:bool":
      return parseOnOff(text) ?? false;
    case "vt:date":
    case "vt:filetime":
      return text ? new Date(text) : undefined;
    default:
      if (NUMERIC_VT_TAGS.has(el.name ?? "")) return text === "" ? undefined : Number(text);
      return text || undefined;
  }
}

/**
 * Serialize a vt:vector of lpstr entries (the TitlesOfParts shape).
 * lpstr is required: Excel refuses to open a file whose app.xml vectors
 * carry lpwstr (CT_VectorLpstr).
 */
export function stringifyStringVector(values: readonly string[]): string {
  const items = values.map((v) => `<vt:lpstr>${escapeXml(v)}</vt:lpstr>`).join("");
  return `<vt:vector size="${values.length}" baseType="lpstr">${items}</vt:vector>`;
}

/** Parse a vt:vector with any scalar baseType into its JS values. */
export function parseVector(el: Element | undefined): VariantValue[] {
  if (!el || el.name !== "vt:vector") return [];
  const out: VariantValue[] = [];
  for (const child of el.elements ?? []) {
    // baseType="variant" wraps each entry in a <vt:variant> carrier.
    const scalar =
      child.name === "vt:variant" ? child.elements?.find((e) => e.name?.startsWith("vt:")) : child;
    if (!scalar) continue;
    const value = parseVariantValue(scalar);
    if (value !== undefined) out.push(value);
  }
  return out;
}
