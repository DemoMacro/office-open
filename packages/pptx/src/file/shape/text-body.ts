import { XmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { xsdTextAnchor } from "@office-open/core";
import { xml } from "@office-open/xml";

import { VerticalAlignment } from "../table/table-cell";
import { Paragraph } from "./paragraph/paragraph";
import type { ParagraphOptions } from "./paragraph/paragraph";
import { TextRun } from "./paragraph/run";

export interface TextBodyOptions {
  readonly text?: string;
  readonly children?: readonly (Paragraph | ParagraphOptions | string)[];
  readonly vertical?: "vert" | "vert270" | "horz" | "wordArtVert";
  readonly anchor?: (typeof VerticalAlignment)[keyof typeof VerticalAlignment];
  readonly autoFit?: "normal" | "shape" | "none";
  readonly wrap?: "square" | "none";
  readonly margins?: {
    readonly top?: number;
    readonly bottom?: number;
    readonly left?: number;
    readonly right?: number;
  };
  readonly columns?: number;
  readonly columnSpacing?: number;
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
 * p:txBody — Text body within a shape.
 * Lazy: stores options, builds XML in toXml().
 */
export class TextBody extends XmlComponent {
  private readonly options: TextBodyOptions;

  public constructor(options: TextBodyOptions = {}) {
    super("p:txBody");
    this.options = options;
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

    return `<p:txBody>${parts.join("")}</p:txBody>`;
  }
}
