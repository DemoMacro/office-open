/**
 * Row-level Structured Document Tag (SDT) for WordprocessingML documents.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SdtRow
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { TableRow } from "../table";
import {
    StructuredDocumentTagContent,
    StructuredDocumentTagProperties,
} from "../table-of-contents";
import type { SdtPropertiesOptions } from "../table-of-contents";
import { TableCell } from "../table/table-cell/table-cell";

/**
 * Options for creating a row-level Structured Document Tag (CT_SdtRow).
 */
export interface ISdtRowOptions {
    readonly properties: SdtPropertiesOptions;
    readonly children?: readonly TableRow[];
}

/**
 * A row-level Structured Document Tag (CT_SdtRow).
 *
 * Represents a content control within a table row. The SDT wraps
 * `TableRow` children in `w:sdtContent`.
 *
 * @example
 * ```typescript
 * new StructuredDocumentTagRow({
 *   properties: { alias: "Controlled Row" },
 *   children: [
 *     new TableRow({ children: [new TableCell({ children: [new Paragraph("Cell")] })] }),
 *   ],
 * });
 * ```
 */
export class StructuredDocumentTagRow extends XmlComponent {
    private readonly rows: readonly TableRow[];

    public constructor(options: ISdtRowOptions) {
        super("w:sdt");
        this.rows = options.children ?? [];
        this.root.push(new StructuredDocumentTagProperties(options.properties));
        if (this.rows.length > 0) {
            const content = new StructuredDocumentTagContent();
            for (const row of this.rows) {
                content.addChildElement(row);
            }
            this.root.push(content);
        }
    }

    public get CellCount(): number {
        return Math.max(...this.rows.map((r) => r.CellCount), 0);
    }

    public get cells(): readonly TableCell[] {
        return this.rows.flatMap((r) => r.cells);
    }

    public addCellToColumnIndex(cell: TableCell, columnIndex: number): void {
        if (this.rows.length > 0) {
            this.rows[0].addCellToColumnIndex(cell, columnIndex);
        }
    }
}
