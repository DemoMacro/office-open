/**
 * Direct XML string builders for table properties.
 *
 * Replaces `buildTableProperties() + xml()`, `buildTableRowProperties() + xml()`,
 * `buildTableCellProperties() + xml()`, and `new TablePropertyExceptions().toXml()`
 * with direct string concatenation — zero IXmlableObject allocation.
 *
 * @module
 */

import { xsdVerticalMergeRev } from "@office-open/core";
import {
  measurementOrPercentValue,
  signedTwipsMeasureValue,
  twipsMeasureValue,
} from "@office-open/core";
import type { AlignmentType } from "@parts/paragraph";
import type { TableCellSpacingProperties } from "@parts/table/table-cell-spacing";
import type {
  TableCellBordersOptions,
  TextDirection,
} from "@parts/table/table-cell/table-cell-components";
import { VerticalMergeType } from "@parts/table/table-cell/table-cell-components";
import type { TableBordersOptions } from "@parts/table/table-properties/table-borders";
import type { TableCellMarginOptions } from "@parts/table/table-properties/table-cell-margin";
import type { TableFloatOptions } from "@parts/table/table-properties/table-float-properties";
import type { TableLayoutType } from "@parts/table/table-properties/table-layout";
import type { TableLookOptions } from "@parts/table/table-properties/table-look";
import type {
  CnfStyleOptions,
  TableRowPropertiesOptionsBase,
} from "@parts/table/table-row/table-row-properties";
import type { TableWidthProperties } from "@parts/table/table-width";
import { WidthType } from "@parts/table/table-width";
import type { BorderOptions } from "@shared/border";
import { BorderStyle } from "@shared/border";
import type { ShadingAttributesProperties } from "@shared/shading";
import type { ICellMergeAttributes } from "@shared/track-revision";
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";
import type { TableVerticalAlign } from "@shared/vertical-align";

import { attrParts, borderStr, onOff, shadingStr } from "../paragraph/stringify";

// ── Table width string ──

function tableWidthStr(name: string, opts: TableWidthProperties): string {
  const type = opts.type ?? WidthType.AUTO;
  let w = opts.size;
  if (type === WidthType.PERCENTAGE && typeof w === "number") {
    w = `${w}%`;
  }
  const a = attrParts({
    "w:w": w !== undefined ? measurementOrPercentValue(w) : undefined,
    "w:type": type,
  });
  return `<${name} ${a}/>`;
}

// ── Cell margin string ──

function cellMarginChildrenStr(opts: TableCellMarginOptions): string {
  const unitType = opts.marginUnitType ?? WidthType.DXA;
  const parts: string[] = [];
  if (opts.top !== undefined)
    parts.push(tableWidthStr("w:top", { size: opts.top, type: unitType }));
  if (opts.left !== undefined)
    parts.push(tableWidthStr("w:left", { size: opts.left, type: unitType }));
  if (opts.bottom !== undefined)
    parts.push(tableWidthStr("w:bottom", { size: opts.bottom, type: unitType }));
  if (opts.right !== undefined)
    parts.push(tableWidthStr("w:right", { size: opts.right, type: unitType }));
  return parts.join("");
}

function cellMarginStr(tag: string, opts: TableCellMarginOptions): string | undefined {
  const inner = cellMarginChildrenStr(opts);
  return inner ? `<${tag}>${inner}</${tag}>` : undefined;
}

// ── Table borders string ──

const DEFAULT_BORDER: BorderOptions = {
  color: "auto",
  size: 4,
  style: BorderStyle.SINGLE,
};

function tableBordersStr(opts: TableBordersOptions): string {
  const parts: string[] = [];
  parts.push(borderStr("w:top", opts.top ?? DEFAULT_BORDER));
  parts.push(borderStr("w:left", opts.left ?? DEFAULT_BORDER));
  parts.push(borderStr("w:bottom", opts.bottom ?? DEFAULT_BORDER));
  parts.push(borderStr("w:right", opts.right ?? DEFAULT_BORDER));
  parts.push(borderStr("w:insideH", opts.insideHorizontal ?? DEFAULT_BORDER));
  parts.push(borderStr("w:insideV", opts.insideVertical ?? DEFAULT_BORDER));
  return `<w:tblBorders>${parts.join("")}</w:tblBorders>`;
}

// ── Cell borders string ──

function cellBordersStr(opts: TableCellBordersOptions): string | undefined {
  const parts: string[] = [];
  if (opts.top) parts.push(borderStr("w:top", opts.top));
  if (opts.start) parts.push(borderStr("w:start", opts.start));
  if (opts.left) parts.push(borderStr("w:left", opts.left));
  if (opts.bottom) parts.push(borderStr("w:bottom", opts.bottom));
  if (opts.end) parts.push(borderStr("w:end", opts.end));
  if (opts.right) parts.push(borderStr("w:right", opts.right));
  if (opts.topLeftToBottomRight) parts.push(borderStr("w:tl2br", opts.topLeftToBottomRight));
  if (opts.topRightToBottomLeft) parts.push(borderStr("w:tr2bl", opts.topRightToBottomLeft));
  return parts.length > 0 ? `<w:tcBorders>${parts.join("")}</w:tcBorders>` : undefined;
}

// ── Float properties string ──

function floatPropertiesStr(opts: TableFloatOptions): string {
  const a = attrParts({
    "w:horzAnchor": opts.horizontalAnchor,
    "w:vertAnchor": opts.verticalAnchor,
    "w:tblpX":
      opts.absoluteHorizontalPosition !== undefined
        ? signedTwipsMeasureValue(opts.absoluteHorizontalPosition)
        : undefined,
    "w:tblpXSpec": opts.relativeHorizontalPosition,
    "w:tblpY":
      opts.absoluteVerticalPosition !== undefined
        ? signedTwipsMeasureValue(opts.absoluteVerticalPosition)
        : undefined,
    "w:tblpYSpec": opts.relativeVerticalPosition,
    "w:bottomFromText":
      opts.bottomFromText !== undefined ? twipsMeasureValue(opts.bottomFromText) : undefined,
    "w:topFromText":
      opts.topFromText !== undefined ? twipsMeasureValue(opts.topFromText) : undefined,
    "w:leftFromText":
      opts.leftFromText !== undefined ? twipsMeasureValue(opts.leftFromText) : undefined,
    "w:rightFromText":
      opts.rightFromText !== undefined ? twipsMeasureValue(opts.rightFromText) : undefined,
  });
  return `<w:tblpPr ${a}/>`;
}

// ── Table look string ──

function tableLookStr(opts: TableLookOptions): string {
  const a = attrParts({
    "w:firstRow": opts.firstRow,
    "w:lastRow": opts.lastRow,
    "w:firstColumn": opts.firstColumn,
    "w:lastColumn": opts.lastColumn,
    "w:noHBand": opts.noHBand,
    "w:noVBand": opts.noVBand,
  });
  return `<w:tblLook ${a}/>`;
}

// ── Change/revision attribute string ──

function changeAttrStr(tag: string, opts: ChangedAttributesProperties): string {
  const a = attrParts({ "w:author": opts.author, "w:date": opts.date, "w:id": opts.id });
  return `<${tag} ${a}/>`;
}

// ── Cell merge revision string ──

function cellMergeStr(opts: ICellMergeAttributes): string {
  const attrs: Record<string, string | number | boolean | undefined> = {
    "w:author": opts.author,
    "w:date": opts.date,
    "w:id": opts.id,
  };
  if (opts.verticalMerge !== undefined) {
    attrs["w:vMerge"] = xsdVerticalMergeRev.to(opts.verticalMerge);
  }
  if (opts.verticalMergeOriginal !== undefined) {
    attrs["w:vMergeOrig"] = xsdVerticalMergeRev.to(opts.verticalMergeOriginal);
  }
  const a = attrParts(attrs);
  return `<w:cellMerge ${a}/>`;
}

// ── Cell spacing string ──

function cellSpacingStr(opts: TableCellSpacingProperties): string {
  const a = attrParts({
    "w:type": opts.type,
    "w:w": measurementOrPercentValue(opts.value),
  });
  return `<w:tblCellSpacing ${a}/>`;
}

// ── Table properties types ──

export interface TablePropertiesOptionsBase {
  width?: TableWidthProperties;
  indent?: TableWidthProperties;
  layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  borders?: TableBordersOptions;
  float?: TableFloatOptions;
  shading?: ShadingAttributesProperties;
  style?: string;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  cellMargin?: TableCellMarginOptions;
  visuallyRightToLeft?: boolean;
  tableLook?: TableLookOptions;
  cellSpacing?: TableCellSpacingProperties;
  styleRowBandSize?: number;
  styleColBandSize?: number;
  caption?: string;
  description?: string;
}

export type ITablePropertiesChangeOptions = ITablePropertiesOptions & ChangedAttributesProperties;

export type ITablePropertiesOptions = {
  revision?: ITablePropertiesChangeOptions;
  includeIfEmpty?: boolean;
} & TablePropertiesOptionsBase;

// ── Table properties change (w:tblPrChange) ──

function stringifyTablePropertiesChangeInner(options: ITablePropertiesChangeOptions): string {
  const inner = stringifyTablePropertiesInner({ ...options, includeIfEmpty: true });
  const a = attrParts({ "w:author": options.author, "w:date": options.date, "w:id": options.id });
  return `<w:tblPrChange ${a}><w:tblPr>${inner}</w:tblPr></w:tblPrChange>`;
}

// ── Table properties (w:tblPr) ──

function stringifyTablePropertiesInner(options: ITablePropertiesOptions): string {
  const parts: string[] = [];

  if (options.style) {
    parts.push(`<w:tblStyle w:val="${options.style}"/>`);
  }

  if (options.float) {
    parts.push(floatPropertiesStr(options.float));
    if (options.float.overlap) {
      parts.push(`<w:tblOverlap w:val="${options.float.overlap}"/>`);
    }
  }

  if (options.visuallyRightToLeft !== undefined) {
    parts.push(onOff("w:bidiVisual", options.visuallyRightToLeft));
  }

  if (options.styleRowBandSize !== undefined) {
    parts.push(`<w:tblStyleRowBandSize w:val="${options.styleRowBandSize}"/>`);
  }

  if (options.styleColBandSize !== undefined) {
    parts.push(`<w:tblStyleColBandSize w:val="${options.styleColBandSize}"/>`);
  }

  if (options.width) {
    parts.push(tableWidthStr("w:tblW", options.width));
  }

  if (options.alignment) {
    parts.push(`<w:jc w:val="${options.alignment}"/>`);
  }

  if (options.indent) {
    parts.push(tableWidthStr("w:tblInd", options.indent));
  }

  if (options.borders) {
    parts.push(tableBordersStr(options.borders));
  }

  if (options.shading) {
    parts.push(shadingStr(options.shading));
  }

  if (options.layout) {
    parts.push(`<w:tblLayout w:type="${options.layout}"/>`);
  }

  if (options.cellMargin) {
    const cm = cellMarginStr("w:tblCellMar", options.cellMargin);
    if (cm) parts.push(cm);
  }

  if (options.tableLook) {
    parts.push(tableLookStr(options.tableLook));
  }

  if (options.cellSpacing) {
    parts.push(cellSpacingStr(options.cellSpacing));
  }

  if (options.caption !== undefined) {
    parts.push(`<w:tblCaption w:val="${options.caption}"/>`);
  }

  if (options.description !== undefined) {
    parts.push(`<w:tblDescription w:val="${options.description}"/>`);
  }

  if (options.revision) {
    parts.push(stringifyTablePropertiesChangeInner(options.revision));
  }

  return parts.join("");
}

export function stringifyTableProperties(options: ITablePropertiesOptions): string | undefined {
  const inner = stringifyTablePropertiesInner(options);
  if (options.includeIfEmpty || inner) {
    return `<w:tblPr>${inner}</w:tblPr>`;
  }
  return undefined;
}

// ── Row properties types ──

export type ITableRowPropertiesChangeOptions = TableRowPropertiesOptionsBase &
  ChangedAttributesProperties;

export type ITableRowPropertiesOptions = TableRowPropertiesOptionsBase & {
  insertion?: ChangedAttributesProperties;
  deletion?: ChangedAttributesProperties;
  revision?: ITableRowPropertiesChangeOptions;
  includeIfEmpty?: boolean;
};

// ── Row properties change (w:trPrChange) ──

function stringifyTableRowPropertiesChangeInner(options: ITableRowPropertiesChangeOptions): string {
  const inner = stringifyTableRowPropertiesInner({ ...options, includeIfEmpty: true });
  const a = attrParts({ "w:author": options.author, "w:date": options.date, "w:id": options.id });
  return `<w:trPrChange ${a}><w:trPr>${inner}</w:trPr></w:trPrChange>`;
}

// ── Row properties (w:trPr) ──

function stringifyTableRowPropertiesInner(options: ITableRowPropertiesOptions): string {
  const parts: string[] = [];

  if (options.cnfStyle !== undefined) {
    const a = attrParts({ "w:val": options.cnfStyle.val, "w:changed": options.cnfStyle.changed });
    parts.push(`<w:cnfStyle ${a}/>`);
  }

  if (options.divId !== undefined) {
    parts.push(`<w:divId w:val="${options.divId}"/>`);
  }

  if (options.gridBefore !== undefined) {
    parts.push(`<w:gridBefore w:val="${options.gridBefore}"/>`);
  }

  if (options.gridAfter !== undefined) {
    parts.push(`<w:gridAfter w:val="${options.gridAfter}"/>`);
  }

  if (options.widthBefore) {
    parts.push(tableWidthStr("w:wBefore", options.widthBefore));
  }

  if (options.widthAfter) {
    parts.push(tableWidthStr("w:wAfter", options.widthAfter));
  }

  if (options.cantSplit !== undefined) {
    parts.push(onOff("w:cantSplit", options.cantSplit));
  }

  if (options.tableHeader !== undefined) {
    parts.push(onOff("w:tblHeader", options.tableHeader));
  }

  if (options.height) {
    const a = attrParts({
      "w:val": twipsMeasureValue(options.height.value),
      "w:hRule": options.height.rule,
    });
    parts.push(`<w:trHeight ${a}/>`);
  }

  if (options.cellSpacing) {
    parts.push(cellSpacingStr(options.cellSpacing));
  }

  if (options.rowAlignment) {
    parts.push(`<w:jc w:val="${options.rowAlignment}"/>`);
  }

  if (options.hidden !== undefined) {
    parts.push(onOff("w:hidden", options.hidden));
  }

  if (options.insertion) {
    parts.push(changeAttrStr("w:ins", options.insertion));
  }

  if (options.deletion) {
    parts.push(changeAttrStr("w:del", options.deletion));
  }

  if (options.revision) {
    parts.push(stringifyTableRowPropertiesChangeInner(options.revision));
  }

  return parts.join("");
}

export function stringifyTableRowProperties(
  options: ITableRowPropertiesOptions,
): string | undefined {
  const inner = stringifyTableRowPropertiesInner(options);
  if (options.includeIfEmpty || inner) {
    return `<w:trPr>${inner}</w:trPr>`;
  }
  return undefined;
}

// ── Cell properties types ──

export interface TableCellPropertiesOptionsBase {
  cnfStyle?: CnfStyleOptions;
  shading?: ShadingAttributesProperties;
  margins?: TableCellMarginOptions;
  verticalAlign?: TableVerticalAlign;
  textDirection?: (typeof TextDirection)[keyof typeof TextDirection];
  verticalMerge?: (typeof VerticalMergeType)[keyof typeof VerticalMergeType];
  width?: TableWidthProperties;
  columnSpan?: number;
  rowSpan?: number;
  borders?: TableCellBordersOptions;
  horizontalMerge?: "continue" | "restart";
  noWrap?: boolean;
  fitText?: boolean;
  hideMark?: boolean;
  headers?: string[];
  insertion?: ChangedAttributesProperties;
  deletion?: ChangedAttributesProperties;
  cellMerge?: ICellMergeAttributes;
}

export type ITableCellPropertiesChangeOptions = TableCellPropertiesOptionsBase &
  ChangedAttributesProperties;

export type ITableCellPropertiesOptions = {
  revision?: ITableCellPropertiesChangeOptions;
  includeIfEmpty?: boolean;
} & TableCellPropertiesOptionsBase;

// ── Cell properties change (w:tcPrChange) ──

function stringifyTableCellPropertiesChangeInner(
  options: ITableCellPropertiesChangeOptions,
): string {
  const inner = stringifyTableCellPropertiesInner({ ...options, includeIfEmpty: true });
  const a = attrParts({ "w:author": options.author, "w:date": options.date, "w:id": options.id });
  return `<w:tcPrChange ${a}><w:tcPr>${inner}</w:tcPr></w:tcPrChange>`;
}

// ── Cell properties (w:tcPr) ──

function stringifyTableCellPropertiesInner(options: ITableCellPropertiesOptions): string {
  const parts: string[] = [];

  if (options.cnfStyle !== undefined) {
    const a = attrParts({ "w:val": options.cnfStyle.val, "w:changed": options.cnfStyle.changed });
    parts.push(`<w:cnfStyle ${a}/>`);
  }

  if (options.width) {
    parts.push(tableWidthStr("w:tcW", options.width));
  }

  if (options.columnSpan) {
    parts.push(`<w:gridSpan w:val="${options.columnSpan}"/>`);
  }

  if (options.verticalMerge) {
    parts.push(`<w:vMerge w:val="${options.verticalMerge}"/>`);
  } else if (options.rowSpan && options.rowSpan > 1) {
    parts.push(`<w:vMerge w:val="${VerticalMergeType.RESTART}"/>`);
  }

  if (options.borders) {
    const bs = cellBordersStr(options.borders);
    if (bs) parts.push(bs);
  }

  if (options.shading) {
    parts.push(shadingStr(options.shading));
  }

  if (options.margins) {
    const cm = cellMarginStr("w:tcMar", options.margins);
    if (cm) parts.push(cm);
  }

  if (options.textDirection) {
    parts.push(`<w:textDirection w:val="${options.textDirection}"/>`);
  }

  if (options.verticalAlign) {
    parts.push(`<w:vAlign w:val="${options.verticalAlign}"/>`);
  }

  if (options.horizontalMerge !== undefined) {
    if (options.horizontalMerge === "restart") {
      parts.push(`<w:hMerge w:val="restart"/>`);
    } else {
      parts.push(`<w:hMerge/>`);
    }
  }

  if (options.noWrap !== undefined) {
    parts.push(onOff("w:noWrap", options.noWrap));
  }

  if (options.fitText !== undefined) {
    parts.push(onOff("w:tcFitText", options.fitText));
  }

  if (options.hideMark !== undefined) {
    parts.push(onOff("w:hideMark", options.hideMark));
  }

  if (options.headers !== undefined) {
    const headerParts = options.headers.map((h) => `<w:header w:val="${h}"/>`).join("");
    parts.push(`<w:headers>${headerParts}</w:headers>`);
  }

  if (options.insertion) {
    parts.push(changeAttrStr("w:cellIns", options.insertion));
  }

  if (options.deletion) {
    parts.push(changeAttrStr("w:cellDel", options.deletion));
  }

  if (options.revision) {
    parts.push(stringifyTableCellPropertiesChangeInner(options.revision));
  }

  if (options.cellMerge) {
    parts.push(cellMergeStr(options.cellMerge));
  }

  return parts.join("");
}

export function stringifyTableCellProperties(
  options: ITableCellPropertiesOptions,
): string | undefined {
  const inner = stringifyTableCellPropertiesInner(options);
  if (options.includeIfEmpty || inner) {
    return `<w:tcPr>${inner}</w:tcPr>`;
  }
  return undefined;
}

// ── Table property exceptions (w:tblPrEx) ──

export interface TablePropertyExOptions {
  width?: TableWidthProperties;
  indent?: TableWidthProperties;
  layout?: (typeof TableLayoutType)[keyof typeof TableLayoutType];
  borders?: TableBordersOptions;
  shading?: ShadingAttributesProperties;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  cellMargin?: TableCellMarginOptions;
  tableLook?: TableLookOptions;
  cellSpacing?: TableCellSpacingProperties;
  tblPrExChange?: { author: string; date?: string; id: string };
}

export function stringifyTablePropertyExceptions(options: TablePropertyExOptions): string {
  const parts: string[] = [];

  if (options.width) {
    parts.push(tableWidthStr("w:tblW", options.width));
  }

  if (options.alignment) {
    parts.push(`<w:jc w:val="${options.alignment}"/>`);
  }

  if (options.cellSpacing) {
    parts.push(cellSpacingStr(options.cellSpacing));
  }

  if (options.indent) {
    parts.push(tableWidthStr("w:tblInd", options.indent));
  }

  if (options.borders) {
    parts.push(tableBordersStr(options.borders));
  }

  if (options.shading) {
    parts.push(shadingStr(options.shading));
  }

  if (options.layout) {
    parts.push(`<w:tblLayout w:type="${options.layout}"/>`);
  }

  if (options.cellMargin) {
    const cm = cellMarginStr("w:tblCellMar", options.cellMargin);
    if (cm) parts.push(cm);
  }

  if (options.tableLook) {
    parts.push(tableLookStr(options.tableLook));
  }

  if (options.tblPrExChange) {
    const change = options.tblPrExChange;
    const attrs: Record<string, string | number | boolean | undefined> = {
      "w:author": change.author,
      "w:id": change.id,
    };
    if (change.date !== undefined) attrs["w:date"] = change.date;
    const a = attrParts(attrs);
    parts.push(`<w:tblPrExChange ${a}/>`);
  }

  return `<w:tblPrEx>${parts.join("")}</w:tblPrEx>`;
}
