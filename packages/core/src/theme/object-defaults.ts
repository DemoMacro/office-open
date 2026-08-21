/**
 * Object defaults (a:objectDefaults / CT_ObjectStyleDefaults) stringify + parse.
 *
 * Each default definition (spDef/lnDef/txDef) holds shape properties, body
 * properties, list style, and an optional shape style. Reuses core descriptors
 * so defaults round-trip through the same machinery as shapes.
 *
 * @module
 */
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "../descriptor";
import { bodyPropertiesDesc, shapePropertiesDesc, textListStyleDesc } from "../drawing";
import { parseShapeStyle, stringifyShapeStyle } from "./style-matrix";
import type { DefaultShapeDefinitionOptions, ObjectDefaultsOptions } from "./theme-options";

function stringifyDefaultShapeDefinition(
  tag: string,
  opts: DefaultShapeDefinitionOptions,
  ctx: WriteContext,
): string {
  // spPr/bodyPr/lstStyle are minOccurs=1 in CT_DefaultShapeDefinition — a
  // present-but-empty section must still emit its (empty) element. The spPr
  // and lstStyle descriptors return inner content only, so the wrapper tags
  // are added here.
  const spPr =
    opts.shapeProperties !== undefined
      ? `<a:spPr>${shapePropertiesDesc.stringify(opts.shapeProperties, ctx) ?? ""}</a:spPr>`
      : "";
  const bodyPr =
    opts.bodyProperties !== undefined
      ? bodyPropertiesDesc.stringify(opts.bodyProperties, ctx) || "<a:bodyPr/>"
      : "";
  const lstStyle =
    opts.listStyle !== undefined
      ? `<a:lstStyle>${textListStyleDesc.stringify(opts.listStyle, ctx) ?? ""}</a:lstStyle>`
      : "";
  const style = opts.shapeStyle ? stringifyShapeStyle(opts.shapeStyle, ctx) : "";
  return `<${tag}>${spPr}${bodyPr}${lstStyle}${style}</${tag}>`;
}

function parseDefaultShapeDefinition(
  el: XmlElement | undefined,
  ctx: ReadContext,
): DefaultShapeDefinitionOptions | undefined {
  if (!el) return undefined;
  const result: Partial<DefaultShapeDefinitionOptions> = {};
  // An empty element still marks the section present (the stringify leg
  // re-emits the empty marker), so `?? {}` — never drop on empty.
  const spPr = findChild(el, "a:spPr");
  if (spPr) result.shapeProperties = shapePropertiesDesc.parse(spPr, ctx) ?? {};
  const bodyPr = findChild(el, "a:bodyPr");
  if (bodyPr) result.bodyProperties = bodyPropertiesDesc.parse(bodyPr, ctx) ?? {};
  const lstStyle = findChild(el, "a:lstStyle");
  if (lstStyle) result.listStyle = textListStyleDesc.parse(lstStyle, ctx) ?? {};
  const style = findChild(el, "a:style");
  if (style) {
    const shapeStyle = parseShapeStyle(style, ctx);
    if (shapeStyle) result.shapeStyle = shapeStyle;
  }
  return Object.keys(result).length > 0 ? (result as DefaultShapeDefinitionOptions) : undefined;
}

/** Serialize a:objectDefaults. Returns an empty element when no defaults set. */
export function stringifyObjectDefaults(
  opts: ObjectDefaultsOptions | undefined,
  ctx: WriteContext,
): string {
  if (!opts) return "<a:objectDefaults/>";
  const spDef = opts.shapeDefault
    ? stringifyDefaultShapeDefinition("a:spDef", opts.shapeDefault, ctx)
    : "";
  const lnDef = opts.lineDefault
    ? stringifyDefaultShapeDefinition("a:lnDef", opts.lineDefault, ctx)
    : "";
  const txDef = opts.textDefault
    ? stringifyDefaultShapeDefinition("a:txDef", opts.textDefault, ctx)
    : "";
  if (!spDef && !lnDef && !txDef) return "<a:objectDefaults/>";
  return `<a:objectDefaults>${spDef}${lnDef}${txDef}</a:objectDefaults>`;
}

/** Parse a:objectDefaults. */
export function parseObjectDefaults(
  el: XmlElement | undefined,
  ctx: ReadContext,
): ObjectDefaultsOptions | undefined {
  if (!el) return undefined;
  const result: Partial<ObjectDefaultsOptions> = {};
  const shapeDefault = parseDefaultShapeDefinition(findChild(el, "a:spDef"), ctx);
  if (shapeDefault) result.shapeDefault = shapeDefault;
  const lineDefault = parseDefaultShapeDefinition(findChild(el, "a:lnDef"), ctx);
  if (lineDefault) result.lineDefault = lineDefault;
  const textDefault = parseDefaultShapeDefinition(findChild(el, "a:txDef"), ctx);
  if (textDefault) result.textDefault = textDefault;
  return Object.keys(result).length > 0 ? (result as ObjectDefaultsOptions) : undefined;
}
