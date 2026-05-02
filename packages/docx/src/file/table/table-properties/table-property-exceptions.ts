/**
 * Table property exceptions module for WordprocessingML documents.
 *
 * Property exceptions allow overriding table properties at the row level.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_TblPrExBase
 *
 * @module
 */
import { IgnoreIfEmptyXmlComponent } from "@file/xml-components";

import { createAlignment } from "../../paragraph";
import type { AlignmentType } from "../../paragraph";
import { createShading } from "../../shading";
import type { IShadingAttributesProperties } from "../../shading";
import { createTableCellSpacing } from "../table-cell-spacing";
import type { ITableCellSpacingProperties } from "../table-cell-spacing";
import { createTableWidthElement } from "../table-width";
import type { ITableWidthProperties } from "../table-width";
import { TableBorders } from "./table-borders";
import type { ITableBordersOptions } from "./table-borders";
import { createTableCellMargin } from "./table-cell-margin";
import type { ITableCellMarginOptions } from "./table-cell-margin";
import { createTableLayout } from "./table-layout";
import type { TableLayoutType } from "./table-layout";
import { createTableLook } from "./table-look";
import type { ITableLookOptions } from "./table-look";

/**
 * Options for table property exceptions (w:tblPrEx).
 *
 * These override the parent table's properties for a specific row.
 * Subset of ITablePropertiesOptionsBase — excludes style, float, bidiVisual,
 * styleRowBandSize, styleColBandSize, caption, description.
 */
export interface ITablePropertyExOptions {
    readonly width?: ITableWidthProperties;
    readonly indent?: ITableWidthProperties;
    readonly layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
    readonly borders?: ITableBordersOptions;
    readonly shading?: IShadingAttributesProperties;
    readonly alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    readonly cellMargin?: ITableCellMarginOptions;
    readonly tableLook?: ITableLookOptions;
    readonly cellSpacing?: ITableCellSpacingProperties;
}

/**
 * Represents table property exceptions (w:tblPrEx) in a WordprocessingML document.
 *
 * Allows overriding table-level properties for a specific row.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_TblPrEx
 */
export class TablePropertyExceptions extends IgnoreIfEmptyXmlComponent {
    public constructor(options: ITablePropertyExOptions) {
        super("w:tblPrEx", true);

        if (options.width) {
            this.root.push(createTableWidthElement("w:tblW", options.width));
        }

        if (options.alignment) {
            this.root.push(createAlignment(options.alignment));
        }

        if (options.cellSpacing) {
            this.root.push(createTableCellSpacing(options.cellSpacing));
        }

        if (options.indent) {
            this.root.push(createTableWidthElement("w:tblInd", options.indent));
        }

        if (options.borders) {
            this.root.push(new TableBorders(options.borders));
        }

        if (options.shading) {
            this.root.push(createShading(options.shading));
        }

        if (options.layout) {
            this.root.push(createTableLayout(options.layout));
        }

        if (options.cellMargin) {
            const cellMargin = createTableCellMargin(options.cellMargin);
            if (cellMargin) {
                this.root.push(cellMargin);
            }
        }

        if (options.tableLook) {
            this.root.push(createTableLook(options.tableLook));
        }
    }
}
