import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { xsdTextAnchor } from "@office-open/core";

import type { FillOptions } from "../drawingml/fill";
import { Paragraph } from "../shape/paragraph/paragraph";
import type { ParagraphOptions } from "../shape/paragraph/paragraph";
import { TextRun } from "../shape/paragraph/run";
import { TableCellProperties, type CellBorderOptions } from "./table-cell-properties";

export type VerticalAlignment = "top" | "center" | "bottom" | "justify" | "distribute";

export interface TableCellOptions {
  readonly text?: string;
  readonly children?: readonly (BaseXmlComponent | ParagraphOptions | string)[];
  readonly fill?: FillOptions;
  readonly borders?: {
    readonly top?: CellBorderOptions;
    readonly bottom?: CellBorderOptions;
    readonly left?: CellBorderOptions;
    readonly right?: CellBorderOptions;
  };
  readonly columnSpan?: number;
  readonly rowSpan?: number;
  readonly verticalAlign?: VerticalAlignment;
  readonly margins?: {
    readonly top?: number;
    readonly bottom?: number;
    readonly left?: number;
    readonly right?: number;
  };
}

/**
 * a:tc — Table cell with text body and properties.
 * Lazy: stores options, builds XML string in toXml.
 */
export class TableCell extends BaseXmlComponent {
  private readonly options: TableCellOptions;
  private readonly paragraphs?: readonly BaseXmlComponent[];

  public constructor(options: TableCellOptions = {}) {
    super("a:tc");
    this.options = options;
    this.paragraphs =
      options.children?.map((c) =>
        typeof c === "string"
          ? new Paragraph(c)
          : c instanceof BaseXmlComponent
            ? c
            : new Paragraph(c),
      ) ??
      (options.text !== undefined
        ? [
            new Paragraph({
              properties: {},
              children: [new TextRun({ text: options.text })],
            }),
          ]
        : undefined);
  }

  public override toXml(context: Context): string {
    const opts = this.options;
    const parts: string[] = [];

    // gridSpan / rowSpan attributes
    const tcAttrs: string[] = [];
    if (opts.columnSpan !== undefined && opts.columnSpan > 1)
      tcAttrs.push(`gridSpan="${opts.columnSpan}"`);
    if (opts.rowSpan !== undefined && opts.rowSpan > 1) tcAttrs.push(`rowSpan="${opts.rowSpan}"`);
    const tcAttrStr = tcAttrs.length > 0 ? ` ${tcAttrs.join(" ")}` : "";

    // a:txBody
    const txParts: string[] = [];
    const margins = opts.margins;
    const bodyPrAttrs: string[] = [];
    if (margins?.top !== undefined) bodyPrAttrs.push(`tIns="${margins.top}"`);
    if (margins?.bottom !== undefined) bodyPrAttrs.push(`bIns="${margins.bottom}"`);
    if (margins?.left !== undefined) bodyPrAttrs.push(`lIns="${margins.left}"`);
    if (margins?.right !== undefined) bodyPrAttrs.push(`rIns="${margins.right}"`);
    const bodyPrStr = bodyPrAttrs.length > 0 ? ` ${bodyPrAttrs.join(" ")}` : "";
    txParts.push(`<a:bodyPr${bodyPrStr}/>`);
    txParts.push("<a:lstStyle/>");

    if (this.paragraphs) {
      for (const p of this.paragraphs) {
        txParts.push(p.toXml(context));
      }
    } else {
      txParts.push("<a:p/>");
    }
    parts.push(`<a:txBody>${txParts.join("")}</a:txBody>`);

    // a:tcPr (has fill needing context, uses toXml() internally)
    const tcPr = new TableCellProperties({
      fill: opts.fill,
      borders: opts.borders,
      verticalAlign: opts.verticalAlign
        ? (xsdTextAnchor.to(opts.verticalAlign) as "t" | "ctr" | "b")
        : undefined,
    });
    parts.push(tcPr.toXml(context));

    return `<a:tc${tcAttrStr}>${parts.join("")}</a:tc>`;
  }
}
