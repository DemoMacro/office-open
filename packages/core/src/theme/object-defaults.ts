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
  const spPrInner = opts.shapeProperties
    ? (shapePropertiesDesc.stringify(opts.shapeProperties, ctx) ?? "")
    : "";
  const spPr = spPrInner ? `<a:spPr>${spPrInner}</a:spPr>` : "";
  const bodyPr = opts.bodyProperties
    ? (bodyPropertiesDesc.stringify(opts.bodyProperties, ctx) ?? "")
    : "";
  const lstStyle = opts.listStyle ? (textListStyleDesc.stringify(opts.listStyle, ctx) ?? "") : "";
  const style = opts.shapeStyle ? stringifyShapeStyle(opts.shapeStyle, ctx) : "";
  return `<${tag}>${spPr}${bodyPr}${lstStyle}${style}</${tag}>`;
}

function parseDefaultShapeDefinition(
  el: XmlElement | undefined,
  ctx: ReadContext,
): DefaultShapeDefinitionOptions | undefined {
  if (!el) return undefined;
  const result: Partial<DefaultShapeDefinitionOptions> = {};
  const spPr = findChild(el, "a:spPr");
  if (spPr) {
    const shapeProperties = shapePropertiesDesc.parse(spPr, ctx);
    if (shapeProperties) result.shapeProperties = shapeProperties;
  }
  const bodyPr = findChild(el, "a:bodyPr");
  if (bodyPr) {
    const bodyProperties = bodyPropertiesDesc.parse(bodyPr, ctx);
    if (bodyProperties) result.bodyProperties = bodyProperties;
  }
  const lstStyle = findChild(el, "a:lstStyle");
  if (lstStyle) {
    const listStyle = textListStyleDesc.parse(lstStyle, ctx);
    if (listStyle) result.listStyle = listStyle;
  }
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
