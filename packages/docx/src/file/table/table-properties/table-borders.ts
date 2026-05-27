/**
 * Table borders module for WordprocessingML documents.
 *
 * This module provides border options for tables.
 *
 * Reference: http://officeopenxml.com/WPtableBorders.php
 *
 * @module
 */
import { BorderStyle, buildBorderObj, createBorderElement } from "@file/border";
import type { BorderOptions } from "@file/border";
import type { IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";

/**
 * Options for configuring table borders.
 *
 * Borders can be applied to the outside edges (top, bottom, left, right)
 * and inside lines (insideHorizontal, insideVertical) of the table.
 */
export interface TableBordersOptions {
  readonly top?: BorderOptions;
  readonly bottom?: BorderOptions;
  readonly left?: BorderOptions;
  readonly right?: BorderOptions;
  readonly insideHorizontal?: BorderOptions;
  readonly insideVertical?: BorderOptions;
}

const NONE_BORDER: BorderOptions = {
  color: "auto",
  size: 0,
  style: BorderStyle.NONE,
};

const DEFAULT_BORDER: BorderOptions = {
  color: "auto",
  size: 4,
  style: BorderStyle.SINGLE,
};

/**
 * Represents table borders in a WordprocessingML document.
 *
 * The tblBorders element specifies the borders for all cells in the table.
 *
 * Reference: http://officeopenxml.com/WPtableBorders.php
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * new TableBorders({
 *   top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
 *   bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
 * });
 *
 * // To remove all borders
 * new TableBorders(TableBorders.NONE);
 * ```
 */
/**
 * Build table borders (w:tblBorders) as IXmlableObject without allocating XmlComponent tree.
 */
export function buildTableBorders(options: TableBordersOptions): IXmlableObject {
  return {
    "w:tblBorders": [
      buildBorderObj("w:top", options.top ?? DEFAULT_BORDER),
      buildBorderObj("w:left", options.left ?? DEFAULT_BORDER),
      buildBorderObj("w:bottom", options.bottom ?? DEFAULT_BORDER),
      buildBorderObj("w:right", options.right ?? DEFAULT_BORDER),
      buildBorderObj("w:insideH", options.insideHorizontal ?? DEFAULT_BORDER),
      buildBorderObj("w:insideV", options.insideVertical ?? DEFAULT_BORDER),
    ],
  };
}

export class TableBorders extends XmlComponent {
  public static readonly NONE: TableBordersOptions = {
    bottom: NONE_BORDER,
    insideHorizontal: NONE_BORDER,
    insideVertical: NONE_BORDER,
    left: NONE_BORDER,
    right: NONE_BORDER,
    top: NONE_BORDER,
  };

  public constructor(options: TableBordersOptions) {
    super("w:tblBorders");

    this.root.push(createBorderElement("w:top", options.top ?? DEFAULT_BORDER));
    this.root.push(createBorderElement("w:left", options.left ?? DEFAULT_BORDER));
    this.root.push(createBorderElement("w:bottom", options.bottom ?? DEFAULT_BORDER));
    this.root.push(createBorderElement("w:right", options.right ?? DEFAULT_BORDER));
    this.root.push(createBorderElement("w:insideH", options.insideHorizontal ?? DEFAULT_BORDER));
    this.root.push(createBorderElement("w:insideV", options.insideVertical ?? DEFAULT_BORDER));
  }
}
