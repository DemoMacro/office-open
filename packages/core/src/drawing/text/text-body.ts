/**
 * Text body descriptor — assembles the inner content of a DrawingML text body
 * (a:bodyPr + a:lstStyle + a:p[]). The caller wraps the container tag
 * (p:txBody / xdr:txBody / a:txBody); DOCX text boxes use wps:txbx (not txBody).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TextBody
 *
 * @module
 */

import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { createBodyProperties, parseBodyProperties } from "./body-properties";
import type { BodyPropertiesOptions } from "./body-properties";
import { textListStyleDesc } from "./list-style";
import type { TextListStyleOptions } from "./list-style";
import { paragraphDesc } from "./paragraph";
import type { ParagraphDescriptorOptions } from "./paragraph";

export interface TextBodyOptions {
  bodyProperties?: BodyPropertiesOptions;
  listStyle?: TextListStyleOptions;
  /** Single-paragraph shorthand; expands to one paragraph on stringify. */
  text?: string;
  /** Paragraph list; string entries expand to a one-run paragraph. */
  paragraphs?: (ParagraphDescriptorOptions | string)[];

  /**
   * Convenience sugar — merged into {@link bodyProperties} on stringify (explicit
   * `bodyProperties` fields take precedence). Input-only: parse always emits
   * the normalized `bodyProperties` form. Mirrors the `fill` descriptor's
   * string-sugar convention.
   */
  anchor?: BodyPropertiesOptions["anchor"];
  /** Sugar → `bodyProperties.normAutofit` (`"normal"`) / `spAutoFit` (`"shape"`). */
  autoFit?: "normal" | "shape";
  /** Sugar → `bodyProperties.numCol`. */
  columns?: BodyPropertiesOptions["numCol"];
  /** Sugar → `bodyProperties.spcCol`. */
  columnSpacing?: BodyPropertiesOptions["spcCol"];
  /** Sugar → `bodyProperties.margins`. */
  margins?: BodyPropertiesOptions["margins"];
  /** Sugar → `bodyProperties.wrap`. */
  wrap?: BodyPropertiesOptions["wrap"];
  /** Sugar → `bodyProperties.vertical`. */
  vertical?: BodyPropertiesOptions["vertical"];
}

/**
 * Collect top-level TextBodyOptions sugar into a partial BodyPropertiesOptions.
 * Returns undefined when no sugar field is set — textBodyDesc runs once per
 * shape and the common case has no sugar, so the guard avoids allocating the
 * collector object (and the caller's emptiness check) on that path.
 */
function bodyPropertiesSugar(opts: TextBodyOptions): Partial<BodyPropertiesOptions> | undefined {
  if (
    opts.anchor === undefined &&
    opts.autoFit === undefined &&
    opts.columns === undefined &&
    opts.columnSpacing === undefined &&
    opts.margins === undefined &&
    opts.wrap === undefined &&
    opts.vertical === undefined
  ) {
    return undefined;
  }
  const sugar: Partial<BodyPropertiesOptions> = {};
  if (opts.anchor !== undefined) sugar.anchor = opts.anchor;
  if (opts.autoFit === "normal") sugar.normAutofit = {};
  else if (opts.autoFit === "shape") sugar.spAutoFit = true;
  if (opts.columns !== undefined) sugar.numCol = opts.columns;
  if (opts.columnSpacing !== undefined) sugar.spcCol = opts.columnSpacing;
  if (opts.margins !== undefined) sugar.margins = opts.margins;
  if (opts.wrap !== undefined) sugar.wrap = opts.wrap;
  if (opts.vertical !== undefined) sugar.vertical = opts.vertical;
  return sugar;
}

export const textBodyDesc: CustomDescriptor<TextBodyOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // a:bodyPr — always present in CT_TextBody. Top-level sugar (anchor/
    // autoFit/columns/...) merges into bodyProperties; explicit bodyProperties
    // fields win.
    const sugar = bodyPropertiesSugar(opts);
    if (opts.bodyProperties || sugar !== undefined) {
      parts.push(createBodyProperties({ ...sugar, ...opts.bodyProperties }));
    } else {
      parts.push("<a:bodyPr/>");
    }

    // a:lstStyle — always emit (matches MS Office byte layout).
    parts.push(
      opts.listStyle
        ? `<a:lstStyle>${textListStyleDesc.stringify(opts.listStyle, ctx) ?? ""}</a:lstStyle>`
        : "<a:lstStyle/>",
    );

    // a:p[] — CT_TextBody requires at least one paragraph. `text` is a
    // single-paragraph shorthand; `paragraphs` accepts string entries that
    // each expand to a one-run paragraph.
    const paragraphs = opts.paragraphs ?? (opts.text !== undefined ? [{ text: opts.text }] : []);
    if (paragraphs.length > 0) {
      for (const p of paragraphs) {
        const para = typeof p === "string" ? { children: [p] } : p;
        parts.push(paragraphDesc.stringify(para, ctx) ?? "<a:p/>");
      }
    } else {
      parts.push("<a:p/>");
    }

    return parts.join("");
  },

  parse(el, ctx) {
    const result: TextBodyOptions = {};

    const bodyPr = findChild(el, "a:bodyPr");
    if (bodyPr) result.bodyProperties = parseBodyProperties(bodyPr, ctx);

    const lstStyle = findChild(el, "a:lstStyle");
    if (lstStyle) {
      const parsed = textListStyleDesc.parse(lstStyle, ctx);
      // An empty <a:lstStyle/> parses to an empty list; skip it so stringify
      // re-emits the self-closing form (matches MS Office byte layout for
      // bare text bodies).
      if (parsed.defaultParagraph || (parsed.levels?.length ?? 0) > 0) {
        result.listStyle = parsed;
      }
    }

    const paragraphs: ParagraphDescriptorOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "a:p") {
        paragraphs.push(paragraphDesc.parse(child, ctx) as ParagraphDescriptorOptions);
      }
    }
    if (paragraphs.length > 0) result.paragraphs = paragraphs;

    return result;
  },
};
