import { BuilderElement, XmlComponent } from "@file/xml-components";

import { Paragraph } from "../shape/paragraph/paragraph";
import { Run } from "../shape/paragraph/run";
import { TableCellProperties } from "./table-cell-properties";
import type { ICellBorderOptions } from "./table-cell-properties";

export interface ITableCellOptions {
    readonly text?: string;
    readonly children?: readonly XmlComponent[];
    readonly fill?: XmlComponent;
    readonly borders?: {
        readonly top?: ICellBorderOptions;
        readonly bottom?: ICellBorderOptions;
        readonly left?: ICellBorderOptions;
        readonly right?: ICellBorderOptions;
    };
    readonly columnSpan?: number;
    readonly rowSpan?: number;
    readonly verticalAlign?: "t" | "ctr" | "b";
    readonly margins?: {
        readonly top?: number;
        readonly bottom?: number;
        readonly left?: number;
        readonly right?: number;
    };
}

/**
 * a:tc — Table cell with text body and properties.
 *
 * Uses a:txBody (DrawingML context), not p:txBody.
 */
export class TableCell extends XmlComponent {
    public constructor(options: ITableCellOptions = {}) {
        super("a:tc");

        // a:txBody (DrawingML namespace, not p:txBody)
        if (options.children) {
            this.root.push(new CellTextBody(options.children, options.margins));
        } else if (options.text !== undefined) {
            this.root.push(
                new CellTextBody([
                    new Paragraph({
                        properties: { bulletNone: false },
                        children: [new Run({ text: options.text })],
                    }),
                ], options.margins),
            );
        } else {
            this.root.push(new CellTextBody(undefined, options.margins));
        }

        // a:tcPr
        this.root.push(
            new TableCellProperties({
                fill: options.fill,
                borders: options.borders,
                columnSpan: options.columnSpan,
                rowSpan: options.rowSpan,
                verticalAlign: options.verticalAlign,
            }),
        );
    }
}

/**
 * a:txBody — Text body for table cells (DrawingML namespace).
 * Differs from shape TextBody (p:txBody) and omits a:buNone.
 */
class CellTextBody extends XmlComponent {
    public constructor(paragraphs?: readonly XmlComponent[], margins?: { readonly top?: number; readonly bottom?: number; readonly left?: number; readonly right?: number }) {
        super("a:txBody");

        const bodyPrAttrs: Record<string, { readonly key: string; readonly value: number }> = {};
        if (margins?.top !== undefined) bodyPrAttrs.tIns = { key: "tIns", value: margins.top };
        if (margins?.bottom !== undefined) bodyPrAttrs.bIns = { key: "bIns", value: margins.bottom };
        if (margins?.left !== undefined) bodyPrAttrs.lIns = { key: "lIns", value: margins.left };
        if (margins?.right !== undefined) bodyPrAttrs.rIns = { key: "rIns", value: margins.right };
        this.root.push(new BuilderElement({ name: "a:bodyPr", attributes: Object.keys(bodyPrAttrs).length > 0 ? bodyPrAttrs : undefined }));
        this.root.push(new BuilderElement({ name: "a:lstStyle" }));

        if (paragraphs) {
            for (const p of paragraphs) {
                this.root.push(p);
            }
        } else {
            this.root.push(new Paragraph({ properties: { bulletNone: false } }));
        }
    }
}
