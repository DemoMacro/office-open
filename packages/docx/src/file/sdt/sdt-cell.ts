/**
 * Cell-level Structured Document Tag (SDT) for WordprocessingML documents.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SdtCell
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { ITableCellOptions, TableCell } from "../table";
import {
    StructuredDocumentTagContent,
    StructuredDocumentTagProperties,
} from "../table-of-contents";
import type { SdtPropertiesOptions } from "../table-of-contents";

/**
 * Options for creating a cell-level Structured Document Tag (CT_SdtCell).
 */
export interface ISdtCellOptions {
    readonly properties: SdtPropertiesOptions;
    readonly children?: readonly TableCell[];
}

/**
 * A cell-level Structured Document Tag (CT_SdtCell).
 *
 * Represents a content control within a table cell. The SDT wraps
 * `TableCell` children in `w:sdtContent`.
 *
 * @example
 * ```typescript
 * new StructuredDocumentTagCell({
 *   properties: { alias: "Controlled Cell" },
 *   children: [new TableCell({ children: [new Paragraph("Cell")] })],
 * });
 * ```
 */
export class StructuredDocumentTagCell extends XmlComponent {
    public readonly options: ITableCellOptions;

    public constructor(sdtOptions: ISdtCellOptions) {
        super("w:sdt");
        this.options = { children: [] };
        this.root.push(new StructuredDocumentTagProperties(sdtOptions.properties));
        if (sdtOptions.children && sdtOptions.children.length > 0) {
            const content = new StructuredDocumentTagContent();
            for (const child of sdtOptions.children) {
                content.addChildElement(child);
            }
            this.root.push(content);
        }
    }
}
