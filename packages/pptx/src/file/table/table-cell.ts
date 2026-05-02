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
            this.root.push(new CellTextBody(options.children));
        } else if (options.text !== undefined) {
            this.root.push(
                new CellTextBody([
                    new Paragraph({
                        properties: { bulletNone: false },
                        children: [new Run({ text: options.text })],
                    }),
                ]),
            );
        } else {
            this.root.push(new CellTextBody());
        }

        // a:tcPr
        this.root.push(
            new TableCellProperties({
                fill: options.fill,
                borders: options.borders,
                columnSpan: options.columnSpan,
                rowSpan: options.rowSpan,
            }),
        );
    }
}

/**
 * a:txBody — Text body for table cells (DrawingML namespace).
 * Differs from shape TextBody (p:txBody) and omits a:buNone.
 */
class CellTextBody extends XmlComponent {
    public constructor(paragraphs?: readonly XmlComponent[]) {
        super("a:txBody");

        this.root.push(new BuilderElement({ name: "a:bodyPr" }));
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
