import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import type { FillOptions } from "../drawingml/fill";
import { Paragraph } from "../shape/paragraph/paragraph";
import { Run } from "../shape/paragraph/run";
import { TableCellProperties, type ICellBorderOptions } from "./table-cell-properties";

export const VerticalAlignment = {
    TOP: "t",
    CENTER: "ctr",
    BOTTOM: "b",
} as const;

export interface ITableCellOptions {
    readonly text?: string;
    readonly children?: readonly BaseXmlComponent[];
    readonly fill?: FillOptions;
    readonly borders?: {
        readonly top?: ICellBorderOptions;
        readonly bottom?: ICellBorderOptions;
        readonly left?: ICellBorderOptions;
        readonly right?: ICellBorderOptions;
    };
    readonly columnSpan?: number;
    readonly rowSpan?: number;
    readonly verticalAlign?: keyof typeof VerticalAlignment;
    readonly margins?: {
        readonly top?: number;
        readonly bottom?: number;
        readonly left?: number;
        readonly right?: number;
    };
}

/**
 * a:tc — Table cell with text body and properties.
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class TableCell extends BaseXmlComponent {
    private readonly options: ITableCellOptions;
    private readonly paragraphs?: readonly BaseXmlComponent[];

    public constructor(options: ITableCellOptions = {}) {
        super("a:tc");
        this.options = options;
        this.paragraphs =
            options.children ??
            (options.text !== undefined
                ? [
                      new Paragraph({
                          properties: { bulletNone: false },
                          children: [new Run({ text: options.text })],
                      }),
                  ]
                : undefined);
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const opts = this.options;
        const children: IXmlableObject[] = [];

        // gridSpan and rowSpan attributes
        const tcAttrs: Record<string, number> = {};
        if (opts.columnSpan !== undefined && opts.columnSpan > 1)
            tcAttrs.gridSpan = opts.columnSpan;
        if (opts.rowSpan !== undefined && opts.rowSpan > 1) tcAttrs.rowSpan = opts.rowSpan;
        if (Object.keys(tcAttrs).length > 0) children.push({ _attr: tcAttrs });

        // a:txBody
        const txBodyChildren: IXmlableObject[] = [];
        const margins = opts.margins;
        const bodyPrAttrs: Record<string, number> = {};
        if (margins?.top !== undefined) bodyPrAttrs.tIns = margins.top;
        if (margins?.bottom !== undefined) bodyPrAttrs.bIns = margins.bottom;
        if (margins?.left !== undefined) bodyPrAttrs.lIns = margins.left;
        if (margins?.right !== undefined) bodyPrAttrs.rIns = margins.right;
        txBodyChildren.push({
            "a:bodyPr": Object.keys(bodyPrAttrs).length > 0 ? { _attr: bodyPrAttrs } : {},
        });
        txBodyChildren.push({ "a:lstStyle": {} });

        if (this.paragraphs) {
            for (const p of this.paragraphs) {
                const pObj = p.prepForXml(context);
                if (pObj) txBodyChildren.push(pObj);
            }
        } else {
            txBodyChildren.push({ "a:p": [] });
        }
        children.push({ "a:txBody": txBodyChildren });

        // a:tcPr
        const tcPr = new TableCellProperties({
            fill: opts.fill,
            borders: opts.borders,
            verticalAlign: opts.verticalAlign ? VerticalAlignment[opts.verticalAlign] : undefined,
        });
        const tcPrObj = tcPr.prepForXml(context);
        if (tcPrObj) children.push(tcPrObj);

        return { "a:tc": children };
    }
}
