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
  paragraphs?: ParagraphDescriptorOptions[];
}

export const textBodyDesc: CustomDescriptor<TextBodyOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // a:bodyPr — always present in CT_TextBody.
    parts.push(opts.bodyProperties ? createBodyProperties(opts.bodyProperties) : "<a:bodyPr/>");

    // a:lstStyle — always emit (matches MS Office byte layout).
    parts.push(
      opts.listStyle
        ? `<a:lstStyle>${textListStyleDesc.stringify(opts.listStyle, ctx) ?? ""}</a:lstStyle>`
        : "<a:lstStyle/>",
    );

    // a:p[] — CT_TextBody requires at least one paragraph.
    if (opts.paragraphs && opts.paragraphs.length > 0) {
      for (const p of opts.paragraphs) {
        parts.push(paragraphDesc.stringify(p, ctx) ?? "<a:p/>");
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
    if (lstStyle) result.listStyle = textListStyleDesc.parse(lstStyle, ctx);

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
