import { XmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { xsdTextAnchor } from "@office-open/core";
import { xml } from "@office-open/xml";

import type { VerticalAlignment } from "../table/table-cell";
import { Paragraph } from "./paragraph/paragraph";
import type { ParagraphOptions } from "./paragraph/paragraph";
import { TextRun } from "./paragraph/run";

export interface TextBodyOptions {
  text?: string;
  children?: (Paragraph | ParagraphOptions | string)[];
  vertical?:
    | "horz"
    | "vert"
    | "vert270"
    | "wordArtVert"
    | "eaVert"
    | "mongolianVert"
    | "wordArtVertRtl";
  anchor?: VerticalAlignment;
  autoFit?: "normal" | "shape" | "none";
  wrap?: "square" | "none";
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  marginTop?: number;
  marginBottom?: number;
  columns?: number;
  columnSpacing?: number;
}

/**
 * Pure function: builds a:bodyPr content.
 */
function buildBodyPr(options: TextBodyOptions): IXmlableObject {
  const bodyPrChildren: IXmlableObject[] = [];

  if (options.autoFit === "normal") bodyPrChildren.push({ "a:normAutofit": {} });
  else if (options.autoFit === "shape") bodyPrChildren.push({ "a:spAutoFit": {} });
  else if (options.autoFit === "none") bodyPrChildren.push({ "a:noAutofit": {} });

  const bodyPrContent: IXmlableObject[] = [];

  const attrs: Record<string, string | number> = {};
  if (options.vertical) attrs.vert = options.vertical;
  if (options.anchor) attrs.anchor = xsdTextAnchor.to(options.anchor);
  if (options.wrap) attrs.wrap = options.wrap;
  if (options.margins?.top !== undefined) attrs.tIns = options.margins.top;
  if (options.margins?.bottom !== undefined) attrs.bIns = options.margins.bottom;
  if (options.margins?.left !== undefined) attrs.lIns = options.margins.left;
  if (options.margins?.right !== undefined) attrs.rIns = options.margins.right;
  if (options.columns !== undefined) attrs.numCol = options.columns;
  if (options.columnSpacing !== undefined) attrs.spcCol = options.columnSpacing * 100;
  if (options.marginTop !== undefined) attrs.marT = options.marginTop;
  if (options.marginBottom !== undefined) attrs.marB = options.marginBottom;
  if (Object.keys(attrs).length > 0) bodyPrContent.push({ _attr: attrs });

  bodyPrContent.push(...bodyPrChildren);

  return {
    "a:bodyPr":
      bodyPrContent.length === 1 && "_attr" in bodyPrContent[0]
        ? bodyPrContent[0]
        : bodyPrContent.length > 0
          ? bodyPrContent
          : {},
  };
}

/**
 * p:txBody / a:txBody — Text body within a shape.
 * Lazy: stores options, builds XML in toXml().
 * @param ns - Namespace prefix for the wrapper tag. Defaults to "p" (PresentationML).
 *             Use "a" for DrawingML contexts (e.g., locked canvas).
 */
export class TextBody extends XmlComponent {
  private options: TextBodyOptions;
  private ns: "p" | "a";

  public constructor(options: TextBodyOptions = {}, ns: "p" | "a" = "p") {
    super(`${ns}:txBody`);
    this.options = options;
    this.ns = ns;
  }

  public override toXml(context: Context): string {
    const parts: string[] = [];

    parts.push(xml(buildBodyPr(this.options)));
    parts.push("<a:lstStyle/>");

    if (this.options.children) {
      for (const p of this.options.children) {
        const para =
          typeof p === "string"
            ? new Paragraph({ children: [new TextRun({ text: p })] })
            : p instanceof Paragraph
              ? p
              : new Paragraph(p);
        parts.push(para.toXml(context));
      }
    } else if (this.options.text !== undefined) {
      parts.push(
        new Paragraph({
          children: [new TextRun({ text: this.options.text })],
        }).toXml(context),
      );
    } else {
      parts.push(new Paragraph().toXml(context));
    }

    return `<${this.ns}:txBody>${parts.join("")}</${this.ns}:txBody>`;
  }
}
