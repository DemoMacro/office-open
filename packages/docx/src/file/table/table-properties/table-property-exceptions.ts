/**
 * Table property exceptions module for WordprocessingML documents.
 *
 * Property exceptions allow overriding table properties at the row level.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_TblPrExBase
 *
 * @module
 */
import { BuilderElement, IgnoreIfEmptyXmlComponent } from "@file/xml-components";

import { createAlignment } from "../../paragraph";
import type { AlignmentType } from "../../paragraph";
import { createShading } from "../../shading";
import type { ShadingAttributesProperties } from "../../shading";
import { createTableCellSpacing } from "../table-cell-spacing";
import type { TableCellSpacingProperties } from "../table-cell-spacing";
import { createTableWidthElement } from "../table-width";
import type { TableWidthProperties } from "../table-width";
import { TableBorders } from "./table-borders";
import type { TableBordersOptions } from "./table-borders";
import { createTableCellMargin } from "./table-cell-margin";
import type { TableCellMarginOptions } from "./table-cell-margin";
import { createTableLayout } from "./table-layout";
import type { TableLayoutType } from "./table-layout";
import { createTableLook } from "./table-look";
import type { TableLookOptions } from "./table-look";

/**
 * Options for table property exceptions (w:tblPrEx).
 *
 * These override the parent table's properties for a specific row.
 * Subset of TablePropertiesOptionsBase — excludes style, float, bidiVisual,
 * styleRowBandSize, styleColBandSize, caption, description.
 */
export interface TablePropertyExOptions {
  readonly width?: TableWidthProperties;
  readonly indent?: TableWidthProperties;
  readonly layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  readonly borders?: TableBordersOptions;
  readonly shading?: ShadingAttributesProperties;
  readonly alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  readonly cellMargin?: TableCellMarginOptions;
  readonly tableLook?: TableLookOptions;
  readonly cellSpacing?: TableCellSpacingProperties;
  /** Table property exceptions change tracking */
  readonly tblPrExChange?: { readonly author: string; readonly date?: string; readonly id: string };
}

/**
 * Represents table property exceptions (w:tblPrEx) in a WordprocessingML document.
 *
 * Allows overriding table-level properties for a specific row.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_TblPrEx
 */
export class TablePropertyExceptions extends IgnoreIfEmptyXmlComponent {
  public constructor(options: TablePropertyExOptions) {
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

    if (options.tblPrExChange) {
      const change = options.tblPrExChange;
      const attrs: { key: string; value: string }[] = [
        { key: "w:author", value: change.author },
        { key: "w:id", value: change.id },
      ];
      if (change.date !== undefined) attrs.push({ key: "w:date", value: change.date });
      this.root.push(new BuilderElement({ name: "w:tblPrExChange", attributes: attrs }));
    }
  }
}
