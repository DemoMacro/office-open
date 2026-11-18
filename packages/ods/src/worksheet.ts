import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  graphicFill,
  graphicOutline,
  isExternalOdfUrl,
  lengthToEmu,
  odfGroupTransform,
  parseOdfGroupTransform,
  pushShapeStyle,
  spreadsheetColumnName,
  spreadsheetColumnNumber,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type { GraphicStyle } from "@office-open/odf";
import type {
  CellOptions,
  CellRangeSourceOptions,
  ColumnOptions,
  CommentOptions,
  DetectiveHighlightedRangeOptions,
  DetectiveOperationOptions,
  DxfOptions,
  FormulaOptions,
  HyperlinkOptions,
  MergeCellOptions,
  RowOptions,
  RichTextOptions,
  RichTextRunOptions,
  RichTextRunPropertiesOptions,
  SelectionOptions,
  PictureOptions,
  StyleOptions,
  TableDetectiveOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { cellDecorationXml } from "./comments";
import { cellGraphicXml, type OdsChartFrame } from "./drawing-chart";
import { OdsGenerateError, OdsParseError } from "./error";
import { odfDateValue } from "./numbering";
import type { OdsCellOptions, OdsRowOptions, OdsWorksheetOptions } from "./options";
import {
  cellShapeXml,
  conditionalFormatsXml,
  drawingGroupXml,
  parseDrawingGroup,
  parseDrawingShape,
  tableProtectionXml,
} from "./spreadsheet";
import {
  addCellStyle,
  addDimensionStyle,
  addTableStyle,
  addTextStyle,
  type DimensionStyle,
} from "./styles";
import { assertElementNames, unknownOdsElement } from "./workbook";

export const COLUMN_CONTAINERS = new Set([
  "table:table-columns",
  "table:table-header-columns",
  "table:table-column-group",
]);

export const ROW_CONTAINERS = new Set([
  "table:table-rows",
  "table:table-header-rows",
  "table:table-row-group",
]);

const MAX_COLUMNS = 16384;
const MAX_ROWS = 1048576;
const MAX_MATERIALIZED_EMPTY_REPEATS = 100;

interface RepeatedRow {
  readonly row: OdsRowOptions;
  readonly rowIndex: number;
  readonly repeat: number;
}

interface MergedRange {
  readonly firstColumn: number;
  readonly firstRow: number;
  readonly lastColumn: number;
  readonly lastRow: number;
}

type DetectiveHighlightedRangeDirection = Extract<
  DetectiveHighlightedRangeOptions,
  { direction: unknown }
>["direction"];

function repeatCount(
  element: Element,
  name: string,
  maximum: number = Number.MAX_SAFE_INTEGER,
): number {
  const raw = attributeNumber(element, name);
  if (raw === undefined) return 1;
  if (!Number.isInteger(raw) || raw < 1 || raw > maximum) {
    throw new OdsParseError(
      "Invalid OpenDocument spreadsheet repeat count",
      "content.xml",
      "office:body/office:spreadsheet/table:table",
      name,
      "invalid-repeat-count",
    );
  }
  return raw;
}

function structuralEqual(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (Array.isArray(left) && Array.isArray(right)) {
    return (
      left.length === right.length &&
      left.every((value, index) => structuralEqual(value, right[index]))
    );
  }
  if (typeof left === "object" && left !== null && typeof right === "object" && right !== null) {
    const leftEntries = Object.entries(left).filter(
      ([name]) => name !== "reference" && name !== "repeat",
    );
    const rightEntries = Object.entries(right).filter(
      ([name]) => name !== "reference" && name !== "repeat",
    );
    if (leftEntries.length !== rightEntries.length) return false;
    return leftEntries.every(([name, value]) => {
      const rightValue = (right as Record<string, unknown>)[name];
      return (
        rightEntries.some(([rightName]) => rightName === name) && structuralEqual(value, rightValue)
      );
    });
  }
  return false;
}

function positionalRow(row: RowOptions, rowNumber: number): RowOptions {
  return {
    ...row,
    rowNumber,
    cells: row.cells?.map((cell) => ({
      ...cell,
      reference: cell.reference?.replace(/\d+$/u, String(rowNumber)),
      ...(cell.covered
        ? {
            covered: {
              ...cell.covered,
              reference: cell.covered.reference?.replace(/\d+$/u, String(rowNumber)),
            },
          }
        : {}),
    })),
  };
}

function rowEqual(left: RowOptions, right: RowOptions, leftRow: number, rightRow: number): boolean {
  return structuralEqual(positionalRow(left, leftRow), positionalRow(right, rightRow));
}

function graphicDisplay(element: Element): "printer" | undefined {
  const value = attributeString(element, "drawooo:display");
  if (value === undefined) return undefined;
  if (value !== "printer") {
    throw unknownOdsElement(
      element,
      "/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/draw:frame",
      "drawooo:display",
      "display target has no canonical mapping",
    );
  }
  return "printer";
}

function groupRepeatedRows(rows: readonly OdsRowOptions[], canShareOutput: boolean): RepeatedRow[] {
  const groups: RepeatedRow[] = [];
  for (const [rowIndex, row] of rows.entries()) {
    const previous = groups[groups.length - 1];
    const previousRowNumber = previous
      ? (previous.row.rowNumber ?? previous.rowIndex + 1)
      : undefined;
    const rowNumber = row.rowNumber ?? rowIndex + 1;
    if (
      canShareOutput &&
      previous &&
      previous.row.repeat === undefined &&
      row.repeat === undefined &&
      previousRowNumber !== undefined &&
      rowNumber === previousRowNumber + 1 &&
      rowEqual(previous.row, row, previousRowNumber, rowNumber)
    ) {
      groups[groups.length - 1] = { ...previous, repeat: previous.repeat + 1 };
      continue;
    }
    groups.push({ row, rowIndex, repeat: row.repeat ?? 1 });
  }
  return groups;
}

function hasCellContent(cell: CellOptions): boolean {
  return Object.keys(cell).some((key) => key !== "reference" && key !== "style");
}

function hasVisibleCellStyle(style: StyleOptions | undefined): boolean {
  const border = style?.border;
  return (
    style?.fill !== undefined ||
    [border?.top, border?.bottom, border?.left, border?.right].some(
      (side) => side?.style !== undefined && side.style !== "none",
    )
  );
}

function isLargeEmptyRepeat(repeat: number, hasData: boolean): boolean {
  return !hasData && repeat > MAX_MATERIALIZED_EMPTY_REPEATS;
}

export function worksheetXml(
  worksheet: OdsWorksheetOptions,
  index: number,
  styles: string[],
  embeddedCharts: OdsChartFrame[],
  localNamedExpressions = "",
  validationNames = new Map<string, string>(),
  dxfs: DxfOptions[] = [],
  formsXml = "",
): string {
  const hyperlinks = new Map((worksheet.hyperlinks ?? []).map((link) => [link.cell, link]));
  const conditionalStyleNames = new Map<number, string>();
  for (const rule of (worksheet.conditionalFormats ?? []).flatMap((format) => format.rules)) {
    const dxfId = rule.dxfId;
    if (dxfId === undefined || conditionalStyleNames.has(dxfId)) continue;
    const dxf = dxfs[dxfId];
    if (!dxf) continue;
    conditionalStyleNames.set(
      dxfId,
      addCellStyle(
        {
          font: dxf.font,
          fill: dxf.fill,
          border: dxf.border,
          numFmt: typeof dxf.numFmt === "string" ? dxf.numFmt : dxf.numFmt?.formatCode,
          alignment: dxf.alignment,
          protection: dxf.protection,
        },
        styles,
      ),
    );
  }
  const tableSheetName = worksheet.name ?? `Sheet${index}`;
  const merges: MergedRange[] = (worksheet.mergeCells ?? []).flatMap((merge) => {
    const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/u.exec(merge.ref);
    if (!match) return [];
    const firstColumn = columnNumber(match[1]!);
    const firstRow = Number(match[2]!);
    const lastColumn = columnNumber(match[3]!);
    const lastRow = Number(match[4]!);
    if (firstColumn > lastColumn || firstRow > lastRow) return [];
    return [{ firstColumn, firstRow, lastColumn, lastRow }];
  });
  const sourceRows = (worksheet.rows ?? []).map((row, rowIndex) => ({
    ...row,
    rowNumber: row.rowNumber ?? rowIndex + 1,
  }));
  const sourceRowNumbers = new Set(sourceRows.map((row) => row.rowNumber!));
  const mergeRows = merges.flatMap((merge) =>
    Array.from(
      { length: merge.lastRow - merge.firstRow + 1 },
      (_, offset) => merge.firstRow + offset,
    ),
  );
  const rowsForOutput: OdsRowOptions[] = [
    ...sourceRows,
    ...mergeRows
      .filter((rowNumber) => !sourceRowNumbers.has(rowNumber))
      .map((rowNumber) => ({ rowNumber })),
  ].sort((left, right) => (left.rowNumber ?? 0) - (right.rowNumber ?? 0));
  const explicitColumns = (worksheet.columns ?? []).flatMap((column) => {
    const count = Math.max(1, (column.max ?? column.min) - column.min + 1);
    const styleName =
      column.width === undefined && !column.hidden && column.bestFit === undefined
        ? undefined
        : addDimensionStyle(
            { width: column.width, hidden: column.hidden, bestFit: column.bestFit },
            styles,
            "column",
          );
    return Array.from({ length: count }, () =>
      xmlElement("table:table-column", { "table:style-name": styleName }),
    );
  });
  let maxRowCells = 0;
  for (const row of rowsForOutput) {
    maxRowCells = Math.max(maxRowCells, row.cells?.length ?? 0);
  }
  const columnCount = Math.max(explicitColumns.length, maxRowCells, 1);
  const columns = concatenate(
    Array.from(
      { length: columnCount },
      (_, column) => explicitColumns[column] ?? xmlElement("table:table-column"),
    ),
  );
  const canShareRepeatedRows =
    merges.length === 0 &&
    !(worksheet.comments?.length ?? 0) &&
    !(worksheet.shapes?.length ?? 0) &&
    validationNames.size === 0;
  const commentsByCell = new Map<string, NonNullable<WorksheetOptions["comments"]>[number]>();
  for (const comment of worksheet.comments ?? []) commentsByCell.set(comment.cell, comment);
  const shapesByCell = new Map<string, NonNullable<WorksheetOptions["shapes"]>[number][]>();
  for (const shape of worksheet.shapes ?? []) {
    if (shape.anchorType === "absolute") continue;
    const reference = `${cellAddress(shape.col, shape.row)}`;
    const cellShapes = shapesByCell.get(reference);
    if (cellShapes) cellShapes.push(shape);
    else shapesByCell.set(reference, [shape]);
  }
  const mergesByRow = new Map<number, MergedRange[]>();
  const mergeStartsByCell = new Map<string, MergedRange>();
  for (const merge of merges) {
    for (let row = merge.firstRow; row <= merge.lastRow; row += 1) {
      const rowMerges = mergesByRow.get(row);
      if (rowMerges) rowMerges.push(merge);
      else mergesByRow.set(row, [merge]);
    }
    mergeStartsByCell.set(`${merge.firstRow}:${merge.firstColumn}`, merge);
  }
  const repeatedRows = groupRepeatedRows(rowsForOutput, canShareRepeatedRows);
  const defaultRowHeight =
    typeof worksheet.sheetFormat === "object" ? worksheet.sheetFormat.defaultRowHeight : undefined;
  const defaultRowStyle =
    defaultRowHeight === undefined
      ? undefined
      : addDimensionStyle({ height: defaultRowHeight }, styles, "row");
  const emptyRow = {
    repeat: 1,
    xml: xmlElement("table:table-row", { "table:style-name": defaultRowStyle }),
  };
  const bodyRows: Array<{ repeat: number; xml: string }> = [];
  let expectedRowNumber = 1;
  for (const { row, rowIndex, repeat } of repeatedRows) {
    const rowNumber = row.rowNumber ?? rowIndex + 1;
    while (expectedRowNumber < rowNumber) {
      bodyRows.push(emptyRow);
      expectedRowNumber += 1;
    }
    const styleName =
      row.height === undefined && !row.hidden && defaultRowStyle === undefined
        ? undefined
        : addDimensionStyle(
            {
              height: typeof row.height === "number" ? row.height : defaultRowHeight,
              hidden: row.hidden,
            },
            styles,
            "row",
          );
    const cells = (row.cells ?? []).map((cell, cellIndex) => ({
      cell,
      column: cell.reference ? columnNumber(cell.reference.replace(/\d+$/u, "")) : cellIndex + 1,
    }));
    for (const merge of mergesByRow.get(rowNumber) ?? []) {
      if (rowNumber < merge.firstRow || rowNumber > merge.lastRow) continue;
      for (let column = merge.firstColumn; column <= merge.lastColumn; column += 1) {
        const reference = `${cellAddress(column, rowNumber)}`;
        if (cells.some((entry) => entry.cell.reference === reference)) continue;
        cells.push({ cell: { reference, covered: { reference } }, column });
      }
    }
    cells.sort((left, right) => left.column - right.column);
    const groupedCellXml: { count: number; xml: string }[] = [];
    for (const { cell, column } of cells) {
      const start = mergeStartsByCell.get(`${rowNumber}:${column}`);
      const xml = cellXml(
        cell,
        styles,
        [
          ...cellDecorationXml(worksheet, cell.reference ?? "", commentsByCell),
          ...cellShapeXml(shapesByCell.get(cell.reference ?? ""), cell.reference ?? "", styles),
        ],
        validationNames.get(`${tableSheetName}:${cell.reference ?? ""}`),
        hyperlinks.get(cell.reference ?? ""),
        start
          ? {
              columnsSpanned: start.lastColumn - start.firstColumn + 1,
              rowsSpanned: start.lastRow - start.firstRow + 1,
            }
          : undefined,
      );
      const previous = groupedCellXml[groupedCellXml.length - 1];
      const repeat = cell.repeat ?? 1;
      if (previous?.xml === xml) previous.count += repeat;
      else groupedCellXml.push({ count: repeat, xml });
    }
    const xml = xmlElement(
      "table:table-row",
      { "table:style-name": styleName },
      groupedCellXml.map(({ count, xml }) => {
        if (count === 1) return xml;
        const tag = xml.startsWith("<table:covered-table-cell")
          ? "table:covered-table-cell"
          : "table:table-cell";
        return xml.replace(
          /^<table:(?:covered-)?table-cell(?=[ />])/u,
          `<${tag} table:number-columns-repeated="${count}"`,
        );
      }),
    );
    const previousRow = bodyRows[bodyRows.length - 1];
    if (previousRow?.xml === xml) previousRow.repeat += repeat;
    else {
      bodyRows.push({
        repeat,
        xml,
      });
    }
    expectedRowNumber = rowNumber + repeat;
  }
  const rows = concatenate(
    bodyRows.length
      ? bodyRows.map(({ xml, repeat }) =>
          repeat === 1
            ? xml
            : xml.replace(
                /^<table:table-row(?=[ />])/u,
                `<table:table-row table:number-rows-repeated="${repeat}"`,
              ),
        )
      : defaultRowStyle
        ? []
        : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])],
  );
  const lastWrittenRow = rowsForOutput.length === 0 ? 0 : expectedRowNumber - 1;
  const defaultRepeat = Math.max(0, MAX_ROWS - lastWrittenRow);
  const defaultRows =
    defaultRowStyle && defaultRepeat > 0
      ? xmlElement(
          "table:table-row",
          { "table:style-name": defaultRowStyle, "table:number-rows-repeated": defaultRepeat },
          [xmlElement("table:table-cell")],
        )
      : "";
  const sheetName = worksheet.name ?? `Sheet${index}`;
  const tableStyleName =
    worksheet.sheetView?.showFormulas !== undefined ||
    worksheet.sheetView?.showRowColHeaders !== undefined ||
    worksheet.state === "hidden"
      ? addTableStyle(
          {
            hidden: worksheet.state === "hidden",
            showFormulas: worksheet.sheetView?.showFormulas,
            showColumnHeaders: worksheet.sheetView?.showRowColHeaders,
            showRowHeaders: worksheet.sheetView?.showRowColHeaders,
          },
          styles,
        )
      : undefined;
  const frames = embeddedCharts
    .filter((entry) => entry.worksheet === sheetName)
    .map((entry) =>
      xmlElement(
        "draw:frame",
        {
          "draw:name": entry.name,
          "svg:x": emuToLength(
            entry.chart.absoluteX ?? entry.chart.colOffset ?? entry.chart.col ?? 0,
          ),
          "svg:y": emuToLength(
            entry.chart.absoluteY ?? entry.chart.rowOffset ?? entry.chart.row ?? 0,
          ),
          "svg:width": emuToLength(entry.chart.extentCx ?? 400000),
          "svg:height": emuToLength(entry.chart.extentCy ?? 300000),
          ...(entry.chart.toCol !== undefined && entry.chart.toRow !== undefined
            ? {
                "table:end-cell-address": cellAddress(entry.chart.toCol, entry.chart.toRow),
                "table:end-x": emuToLength(entry.chart.toColOffset ?? 0),
                "table:end-y": emuToLength(entry.chart.toRowOffset ?? 0),
              }
            : {}),
        },
        [
          xmlElement("draw:object", {
            "xlink:href": entry.path ?? entry.name,
            "xlink:type": "simple",
            "xlink:show": "embed",
            "xlink:actuate": "onLoad",
          }),
        ],
      ),
    );
  const imageFrames = (worksheet.images ?? []).map((image, imageIndex) =>
    worksheetImageXml(image, index, imageIndex, styles),
  );
  const topLevelShapes = (worksheet.shapes ?? [])
    .filter((shape) => shape.anchorType === "absolute")
    .flatMap((shape) =>
      cellShapeXml(
        [{ ...shape, colOffset: shape.absoluteX, rowOffset: shape.absoluteY }],
        "A1",
        styles,
      ),
    );
  const selectionXml = worksheet.selection?.length ? gnumericSelectionXml(worksheet.selection) : [];
  const tableSourceXml = worksheet.tableSource
    ? xmlElement("table:table-source", {
        "table:mode": worksheet.tableSource.mode,
        "table:table-name": worksheet.tableSource.tableName,
        "xlink:type": "simple",
        "xlink:href": worksheet.tableSource.href,
        "xlink:actuate": worksheet.tableSource.actuate,
        "table:filter-name": worksheet.tableSource.filterName,
        "table:filter-options": worksheet.tableSource.filterOptions,
        "table:refresh-delay": worksheet.tableSource.refreshDelay,
      })
    : "";
  const scenarioXml = worksheet.scenario
    ? xmlElement("table:scenario", {
        "table:scenario-ranges": worksheet.scenario.ranges,
        "table:is-active": worksheet.scenario.active,
        "table:display-border": worksheet.scenario.displayBorder,
        "table:border-color": worksheet.scenario.borderColor,
        "table:copy-back": worksheet.scenario.copyBack,
        "table:copy-styles": worksheet.scenario.copyStyles,
        "table:copy-formulas": worksheet.scenario.copyFormulas,
        "table:comment": worksheet.scenario.comment,
        "table:protected": worksheet.scenario.protected,
      })
    : "";
  return xmlElement(
    "table:table",
    { "table:name": sheetName, "table:style-name": tableStyleName },
    [
      worksheet.tableTitle === undefined
        ? ""
        : xmlElement("table:title", undefined, [escapeText(worksheet.tableTitle)]),
      worksheet.tableDescription === undefined
        ? ""
        : xmlElement("table:desc", undefined, [escapeText(worksheet.tableDescription)]),
      tableSourceXml,
      scenarioXml,
      formsXml,
      ...(frames.length || imageFrames.length || topLevelShapes.length || worksheet.groups?.length
        ? [
            xmlElement("table:shapes", undefined, [
              ...frames,
              ...imageFrames,
              ...topLevelShapes,
              ...(worksheet.groups ?? []).map((group) => drawingGroupXml(group, styles)),
            ]),
          ]
        : []),
      tableProtectionXml(worksheet),
      columns,
      `${rows}${defaultRows}`,
      conditionalFormatsXml(worksheet.conditionalFormats ?? [], sheetName, conditionalStyleNames),
      localNamedExpressions,
      ...selectionXml,
    ],
  );
}

export function worksheetImagePath(
  image: PictureOptions,
  worksheetIndex: number,
  imageIndex: number,
): string {
  return image.sourcePath ?? `Pictures/sheet${worksheetIndex}-image${imageIndex + 1}.${image.type}`;
}

function worksheetImageXml(
  image: PictureOptions,
  worksheetIndex: number,
  imageIndex: number,
  styles: string[],
): string {
  const path = worksheetImagePath(image, worksheetIndex, imageIndex);
  const x = image.anchorType === "absolute" ? image.absoluteX : image.colOffset;
  const y = image.anchorType === "absolute" ? image.absoluteY : image.rowOffset;
  const attributes = {
    "draw:name": image.name,
    "draw:style-name": pushShapeStyle(image.properties?.fill, image.properties?.outline, styles),
    "svg:x": emuToLength(typeof x === "number" ? x : 0),
    "svg:y": emuToLength(typeof y === "number" ? y : 0),
    "svg:width": emuToLength(
      typeof image.extentCx === "number" ? image.extentCx : (image.properties?.width ?? 0),
    ),
    "svg:height": emuToLength(
      typeof image.extentCy === "number" ? image.extentCy : (image.properties?.height ?? 0),
    ),
    "table:end-cell-address":
      image.toCol !== undefined && image.toRow !== undefined
        ? cellAddress(image.toCol, image.toRow)
        : undefined,
    "table:end-x":
      typeof image.toColOffset === "number" ? emuToLength(image.toColOffset) : image.toColOffset,
    "table:end-y":
      typeof image.toRowOffset === "number" ? emuToLength(image.toRowOffset) : image.toRowOffset,
    "draw:transform": odfGroupTransform(image.properties?.rotation),
  };
  return xmlElement("draw:frame", attributes, [
    xmlElement("draw:image", {
      "xlink:href": image.sourceUrl && image.data === undefined ? image.sourceUrl : path,
      "xlink:type": "simple",
      "xlink:show": "embed",
      "xlink:actuate": "onLoad",
    }),
    ...(image.title ? [xmlElement("svg:title", undefined, [escapeText(image.title)])] : []),
    ...(image.description
      ? [xmlElement("svg:desc", undefined, [escapeText(image.description)])]
      : []),
  ]);
}

function parseWorksheetImages(
  table: Element,
  graphicStyles: Map<string, GraphicStyle>,
  binaries: Record<string, Uint8Array>,
  usedBinaryPaths: Set<string>,
): PictureOptions[] {
  const shapes = childNamed(table, "table:shapes");
  const frames = [...childrenNamed(table, "draw:frame"), ...childrenNamed(shapes, "draw:frame")];
  return frames.flatMap((frame) => {
    const image = childNamed(frame, "draw:image");
    if (image === undefined || childNamed(frame, "draw:object") !== undefined) return [];
    const href = attributeString(image, "xlink:href");
    if (href === undefined) {
      throw new OdsParseError(
        "ODS image frame has no xlink:href",
        "content.xml",
        "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:image",
        "xlink:href",
        "missing-image-reference",
      );
    }
    const type = worksheetImageType(href);
    const transform = parseOdfGroupTransform(attributeString(frame, "draw:transform"));
    if (transform === null)
      throw new OdsParseError(
        "ODS image transform cannot be projected to the canonical picture model",
        "content.xml",
        "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame",
        "draw:transform",
        "unsupported-image-transform",
      );
    const path = isExternalOdfUrl(href) ? undefined : odsPackagePath(href);
    const data = path === undefined ? undefined : binaries[path];
    if (path !== undefined && data === undefined)
      throw new OdsParseError(
        `ODS image package member is missing: ${path}`,
        "content.xml",
        "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:image",
        "draw:image",
        "missing-package-part",
      );
    if (path !== undefined) usedBinaryPaths.add(path);
    const x = (lengthToEmu(attributeString(frame, "svg:x")) ?? 0) + (transform?.x ?? 0);
    const y = (lengthToEmu(attributeString(frame, "svg:y")) ?? 0) + (transform?.y ?? 0);
    const width = lengthToEmu(attributeString(frame, "svg:width")) ?? 0;
    const height = lengthToEmu(attributeString(frame, "svg:height")) ?? 0;
    const graphic = graphicStyles.get(attributeString(frame, "draw:style-name") ?? "");
    const fill = graphicFill(graphic);
    const outline = graphicOutline(graphic);
    const end = parseEndCell(frame);
    return [
      {
        type,
        ...(data === undefined ? {} : { data }),
        ...(href && isExternalOdfUrl(href) ? { sourceUrl: href } : {}),
        ...(path === undefined ? {} : { sourcePath: path }),
        name: attributeString(frame, "draw:name"),
        title: textOf(childNamed(frame, "svg:title")) || undefined,
        description: textOf(childNamed(frame, "svg:desc")) || undefined,
        anchorType: end === undefined ? ("absolute" as const) : ("twoCell" as const),
        col: 1,
        row: 1,
        ...(end === undefined
          ? {
              absoluteX: x,
              absoluteY: y,
              extentCx: width,
              extentCy: height,
            }
          : { colOffset: x, rowOffset: y, ...end }),
        ...(fill || outline || transform?.rotation !== undefined || width || height
          ? {
              properties: {
                ...(width ? { width } : {}),
                ...(height ? { height } : {}),
                ...(fill ? { fill } : {}),
                ...(outline ? { outline } : {}),
                ...(transform?.rotation === undefined ? {} : { rotation: transform.rotation }),
              },
            }
          : {}),
      },
    ];
  });
}

function parseEndCell(
  frame: Element,
): Pick<PictureOptions, "toCol" | "toRow" | "toColOffset" | "toRowOffset"> | undefined {
  const match = /^([A-Z]+)([0-9]+)$/u.exec(attributeString(frame, "table:end-cell-address") ?? "");
  if (!match) return undefined;
  return {
    toCol: columnNumber(match[1]!),
    toRow: Number(match[2]),
    toColOffset: lengthToEmu(attributeString(frame, "table:end-x")),
    toRowOffset: lengthToEmu(attributeString(frame, "table:end-y")),
  };
}

function worksheetImageType(href: string): PictureOptions["type"] {
  const extension = href.split(/[?#]/u, 1)[0]?.split(".").pop()?.toLowerCase();
  if (extension === "jpeg") return "jpg";
  if (
    extension === "png" ||
    extension === "jpg" ||
    extension === "gif" ||
    extension === "webp" ||
    extension === "wmf" ||
    extension === "emf"
  )
    return extension;
  throw new OdsParseError(
    `ODS image type cannot be projected to WorkbookOptions: ${extension ?? "unknown"}`,
    "content.xml",
    "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:image",
    "draw:image",
    "unsupported-image-type",
  );
}

export function odsPackagePath(href: string): string {
  let decoded: string;
  try {
    decoded = decodeURIComponent(href.split(/[?#]/u, 1)[0] ?? "");
  } catch (cause) {
    throw new OdsParseError(
      "ODS image reference is not a valid package URI",
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:image",
      "xlink:href",
      "invalid-package-uri",
      { cause },
    );
  }
  const segments: string[] = [];
  for (const segment of decoded.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      if (segments.length === 0)
        throw new OdsParseError(
          "ODS image reference escapes the package root",
          "content.xml",
          "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:image",
          "xlink:href",
          "invalid-package-uri",
        );
      segments.pop();
    } else segments.push(segment);
  }
  const path = segments.join("/");
  if (path === "")
    throw new OdsParseError(
      "ODS image reference does not identify a package member",
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:image",
      "xlink:href",
      "invalid-package-uri",
    );
  return path;
}

function concatenate(parts: readonly string[]): string {
  return parts.join("");
}

export function cellAddress(column: number, row: number): string {
  return `${spreadsheetColumnName(column)}${row}`;
}

export function columnNumber(column: string): number {
  return spreadsheetColumnNumber(column);
}

export function cellXml(
  cell: CellOptions,
  styles: string[],
  decorations: string[] = [],
  contentValidationName?: string,
  hyperlink?: HyperlinkOptions,
  span?: { columnsSpanned: number; rowsSpanned: number },
): string {
  if (cell.covered) {
    const children = [
      ...(cell.covered.text === undefined
        ? []
        : [xmlElement("text:p", undefined, [escapeText(cell.covered.text)])]),
      ...(cell.covered.graphics ?? []).map((graphic) => cellGraphicXml(graphic)),
    ];
    return xmlElement(
      "table:covered-table-cell",
      {
        "table:style-name":
          typeof cell.style === "object" ? addCellStyle(cell.style, styles) : undefined,
        "office:value-type": cell.covered.valueType,
        "table:content-validation-name": contentValidationName,
      },
      children,
    );
  }
  const formula: FormulaOptions | undefined =
    typeof cell.formula === "string" ? { formula: cell.formula } : cell.formula;
  if (formula !== undefined) assertFormulaRepresentable(formula, cell.reference);
  const cached = cacheAttributes(cell.value);
  const styleOptions = typeof cell.style === "object" ? cell.style : undefined;
  const valueKind = numFmtKind(styleOptions?.numFmt);
  let type = cell.error === undefined ? cached.type : "error";
  if (cell.error === undefined) {
    if (valueKind === "time") type = "time";
    else if (valueKind === "currency") type = "currency";
  }
  const children = [
    ...(cell.cellRangeSource ? [cellRangeSourceXml(cell.cellRangeSource)] : []),
    ...decorations,
    ...(cell.detective ? [tableDetectiveXml(cell.detective)] : []),
    ...(cached.type === "string"
      ? isRichText(cell.value)
        ? richTextXml(cell.value, styles, hyperlink)
        : String(cached.value ?? "")
            .split("\n")
            .map((paragraph) =>
              xmlElement(
                "text:p",
                undefined,
                hyperlinkXml(hyperlink, odfTextNodes(hyperlink?.display ?? paragraph)),
              ),
            )
      : []),
    ...(cell.graphics ?? []).map((graphic) => cellGraphicXml(graphic)),
  ];
  return xmlElement(
    "table:table-cell",
    {
      "table:style-name": styleOptions ? addCellStyle(styleOptions, styles) : undefined,
      "table:number-columns-spanned": span?.columnsSpanned,
      "table:number-rows-spanned": span?.rowsSpanned,
      "table:content-validation-name": contentValidationName,
      "table:formula": formula ? `of:=${formula.formula}` : undefined,
      ...(formula?.type === "array" && formula.reference
        ? matrixSpanAttributes(formula.reference)
        : {}),
      "office:value-type": type,
      "office:value": cached.type === "float" ? cached.value : undefined,
      "office:currency":
        valueKind === "currency"
          ? /^\[\$([^-\]]+)/u.exec(styleOptions?.numFmt ?? "")?.[1]
          : undefined,
      "office:time-value":
        type === "time" ? serialToDuration(Number(cached.value ?? 0)) : undefined,
      "office:boolean-value": cached.type === "boolean" ? cached.value : undefined,
      "office:date-value": cached.type === "date" ? cached.value : undefined,
      "office:string-value": type === "error" ? cell.error : undefined,
    },
    children,
  );
}

function assertFormulaRepresentable(formula: FormulaOptions, reference: string | undefined): void {
  if (formula.type === "array" && formula.reference === undefined) {
    throw new OdsGenerateError(
      "ODS array formulas require a reference range",
      "content.xml",
      `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell${reference ? `[@table:address="${reference}"]` : ""}`,
      "table:formula",
      "array-formula-reference-missing",
    );
  }
  if (formula.type !== undefined && formula.type !== "normal" && formula.type !== "array") {
    throw new OdsGenerateError(
      `ODS cannot represent formula type: ${formula.type}`,
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/table:table/table:table-cell",
      "table:formula",
      "unsupported-formula-type",
    );
  }
  const unsupported = [
    "sharedIndex",
    "aca",
    "dt2D",
    "dtr",
    "del1",
    "del2",
    "inputCell1",
    "inputCell2",
    "calculateCell",
    "arrayContext",
  ] as const;
  const property = unsupported.find((name) => formula[name] !== undefined);
  if (property !== undefined) {
    throw new OdsGenerateError(
      `ODS cannot represent formula property: ${property}`,
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/table:table/table:table-cell",
      "table:formula",
      "unsupported-formula-property",
    );
  }
}

function matrixSpanAttributes(reference: string): {
  "table:number-matrix-columns-spanned": number;
  "table:number-matrix-rows-spanned": number;
} {
  const [start, end = start] = reference.replace(/^\$/, "").split(":");
  const startMatch = /^([A-Z]+)([0-9]+)$/u.exec(start ?? "");
  const endMatch = /^([A-Z]+)([0-9]+)$/u.exec(end ?? "");
  if (!startMatch || !endMatch)
    return { "table:number-matrix-columns-spanned": 1, "table:number-matrix-rows-spanned": 1 };
  const firstColumn = columnNumber(startMatch[1]!);
  const lastColumn = columnNumber(endMatch[1]!);
  const firstRow = Number(startMatch[2]);
  const lastRow = Number(endMatch[2]);
  return {
    "table:number-matrix-columns-spanned": Math.abs(lastColumn - firstColumn) + 1,
    "table:number-matrix-rows-spanned": Math.abs(lastRow - firstRow) + 1,
  };
}

function cellRangeSourceXml(source: CellRangeSourceOptions): string {
  return xmlElement("table:cell-range-source", {
    "table:name": source.name,
    "table:last-column-spanned": source.lastColumnSpanned,
    "table:last-row-spanned": source.lastRowSpanned,
    "xlink:type": "simple",
    "xlink:href": source.href,
    "xlink:actuate": source.actuate,
    "table:filter-name": source.filterName,
    "table:filter-options": source.filterOptions,
    "table:refresh-delay": source.refreshDelay,
  });
}

const HIGHLIGHT_DIRECTIONS: Record<DetectiveHighlightedRangeDirection, string> = {
  fromAnotherTable: "from-another-table",
  toAnotherTable: "to-another-table",
  fromSameTable: "from-same-table",
};

const OPERATION_NAMES: Record<DetectiveOperationOptions["name"], string> = {
  traceDependents: "trace-dependents",
  removeDependents: "remove-dependents",
  tracePrecedents: "trace-precedents",
  removePrecedents: "remove-precedents",
  traceErrors: "trace-errors",
};

function tableDetectiveXml(detective: TableDetectiveOptions): string {
  return xmlElement("table:detective", undefined, [
    ...(detective.highlightedRanges ?? []).map((range) =>
      xmlElement(
        "table:highlighted-range",
        "markedInvalid" in range
          ? { "table:marked-invalid": range.markedInvalid }
          : {
              "table:cell-range-address": range.cellRangeAddress,
              "table:direction": HIGHLIGHT_DIRECTIONS[range.direction],
              "table:contains-error": range.containsError,
            },
      ),
    ),
    ...(detective.operations ?? []).map((operation) =>
      xmlElement("table:operation", {
        "table:name": OPERATION_NAMES[operation.name],
        "table:index": operation.index,
      }),
    ),
  ]);
}

export function cacheAttributes(value: CellOptions["value"]): {
  type?: string;
  value?: string | number | boolean;
} {
  if (typeof value === "number") return { type: "float", value };
  if (typeof value === "boolean") return { type: "boolean", value };
  if (value instanceof Date) return { type: "date", value: odfDateValue(value) };
  if (typeof value === "string") return { type: "string", value };
  if (isRichText(value)) {
    return {
      type: "string",
      value: value.text ?? value.runs?.map((run) => run.text).join("") ?? "",
    };
  }
  return {};
}

function isRichText(value: CellOptions["value"]): value is RichTextOptions {
  return typeof value === "object" && value !== null && !(value instanceof Date);
}

function richTextXml(
  value: RichTextOptions,
  styles: string[],
  hyperlink?: HyperlinkOptions,
): string[] {
  const runs = value.runs ?? (value.text === undefined ? [] : [{ text: value.text }]);
  const paragraphs: string[][] = [[]];
  for (const run of runs) {
    const parts = run.text.split("\n");
    for (const [index, part] of parts.entries()) {
      if (part.length > 0) paragraphs[paragraphs.length - 1]!.push(richRunXml(run, part, styles));
      if (index < parts.length - 1) paragraphs.push([]);
    }
  }
  return paragraphs.map((paragraph) =>
    xmlElement("text:p", undefined, hyperlinkXml(hyperlink, paragraph)),
  );
}

function hyperlinkXml(hyperlink: HyperlinkOptions | undefined, children: string[]): string[] {
  const href = hyperlink?.url ?? (hyperlink?.location ? `#${hyperlink.location}` : undefined);
  return href === undefined
    ? children
    : [xmlElement("text:a", { "xlink:href": href, "xlink:type": "simple" }, children)];
}

function richRunXml(run: RichTextRunOptions, text: string, styles: string[]): string {
  const styleName = run.properties ? addTextStyle(run.properties, styles) : undefined;
  return styleName
    ? xmlElement("text:span", { "text:style-name": styleName }, odfTextNodes(text))
    : odfTextNodes(text).join("");
}

function odfTextNodes(text: string): string[] {
  return text.split(/( +)/u).flatMap((part) => {
    if (!part) return [];
    if (!/^ +$/u.test(part)) return [escapeText(part)];
    return [xmlElement("text:s", part.length === 1 ? undefined : { "text:c": part.length })];
  });
}

function styleWithNumFmt(style: CellOptions["style"], numFmt: string): CellOptions["style"] {
  if (typeof style === "number") return style;
  return { ...style, numFmt: style?.numFmt ?? numFmt };
}

function durationToSerial(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const match = /^(-?)P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?)?$/u.exec(value);
  if (!match) return undefined;
  const sign = match[1] === "-" ? -1 : 1;
  const days = Number(match[2] ?? 0);
  const hours = Number(match[3] ?? 0);
  const minutes = Number(match[4] ?? 0);
  const seconds = Number(match[5] ?? 0);
  return sign * (days + hours / 24 + minutes / 1440 + seconds / 86400);
}

function serialToDuration(value: number): string {
  const negative = value < 0;
  let remaining = Math.abs(value) * 86400;
  const hours = Math.floor(remaining / 3600);
  remaining -= hours * 3600;
  const minutes = Math.floor(remaining / 60);
  const seconds = remaining - minutes * 60;
  const secondText = Number.isInteger(seconds) ? String(seconds) : seconds.toFixed(6);
  return `${negative ? "-" : ""}PT${hours}H${minutes}M${secondText}S`;
}

function numFmtKind(numFmt: string | undefined): "currency" | "time" | undefined {
  if (numFmt === undefined) return undefined;
  if (/^\[\$([^-\]]+)(?:-[^\]]+)?\]/u.test(numFmt)) return "currency";
  if (/\[?h+\]?[^ymd]*m+[^ymd]*(?:s|$)/iu.test(numFmt)) return "time";
  return undefined;
}

export function worksheet(
  table: Element,
  index: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  graphicStyles: Map<string, GraphicStyle> = new Map(),
  textStyles: Map<string, RichTextRunPropertiesOptions> = new Map(),
  binaries: Record<string, Uint8Array> = {},
  usedBinaryPaths: Set<string> = new Set(),
): OdsWorksheetOptions {
  const comments: CommentOptions[] = [];
  const hyperlinks: HyperlinkOptions[] = [];
  const mergeCells: MergeCellOptions[] = [];
  const shapes: WorksheetOptions["shapes"] = [];
  const groups: WorksheetOptions["groups"] = [];
  const tableShapes = childNamed(table, "table:shapes");
  const images = parseWorksheetImages(table, graphicStyles, binaries, usedBinaryPaths);
  for (const shape of tableShapes?.elements ?? []) {
    if (shape.name === "draw:g") {
      groups.push(parseDrawingGroup(shape, 1, 1, graphicStyles, textStyles, true));
    } else if (
      [
        "draw:rect",
        "draw:control",
        "draw:ellipse",
        "draw:circle",
        "draw:line",
        "draw:measure",
        "draw:polygon",
        "draw:polyline",
        "draw:custom-shape",
      ].includes(shape.name ?? "")
    ) {
      shapes.push(parseDrawingShape(shape, 1, 1, graphicStyles, textStyles, true));
    }
  }
  validateTableChildren(table);
  const columns: ColumnOptions[] = [];
  let columnNumber = 1;
  for (const column of tableColumns(table)) {
    const repeat = repeatCount(column, "table:number-columns-repeated");
    const parsed = parseColumn(column, columnNumber, dimensions);
    const remainingColumns = MAX_COLUMNS - columnNumber + 1;
    if (remainingColumns > 0) {
      const visibleRepeat = Math.min(repeat, remainingColumns);
      parsed.max = columnNumber + visibleRepeat - 1;
      const previous = columns[columns.length - 1];
      if (
        previous &&
        previous.max === columnNumber - 1 &&
        previous.width === parsed.width &&
        previous.bestFit === parsed.bestFit &&
        previous.hidden === parsed.hidden
      ) {
        previous.max = parsed.max;
      } else {
        columns.push(parsed);
      }
    }
    columnNumber += repeat;
  }
  const parsedRows = parseRows(
    table,
    dimensions,
    cellStyles,
    graphicStyles,
    comments,
    mergeCells,
    shapes,
    textStyles,
    hyperlinks,
    groups,
  );
  const tableStyle = dimensions.get(attributeString(table, "table:style-name") ?? "");
  const sheetView =
    tableStyle?.showFormulas !== undefined || tableStyle?.showColumnHeaders !== undefined
      ? {
          ...(tableStyle?.showFormulas !== undefined
            ? { showFormulas: tableStyle.showFormulas }
            : {}),
          ...(tableStyle?.showColumnHeaders !== undefined
            ? { showRowColHeaders: tableStyle.showColumnHeaders }
            : {}),
        }
      : undefined;
  const worksheetOptions: WorksheetOptions = {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    ...(columns.length ? { columns } : {}),
    ...(sheetView ? { sheetView } : {}),
    ...(tableStyle?.hidden ? { state: "hidden" as const } : {}),
    ...(parsedRows.defaultRowHeight === undefined
      ? {}
      : { sheetFormat: { defaultRowHeight: parsedRows.defaultRowHeight } }),
    rows: parsedRows.rows,
  };
  const selection = gnumericSelection(table);
  return {
    ...worksheetOptions,
    ...parseTableFeatures(table),
    ...(selection ? { selection } : {}),
    ...(comments.length ? { comments } : {}),
    ...(mergeCells.length ? { mergeCells } : {}),
    ...(shapes?.length ? { shapes } : {}),
    ...(groups.length ? { groups } : {}),
    ...(images.length ? { images } : {}),
    ...(hyperlinks.length ? { hyperlinks } : {}),
  };
}

function parseTableFeatures(
  table: Element,
): Pick<OdsWorksheetOptions, "tableTitle" | "tableDescription" | "scenario" | "tableSource"> {
  const scenario = childNamed(table, "table:scenario");
  const source = childNamed(table, "table:table-source");
  const mode = attributeString(source, "table:mode");
  if (mode !== undefined && mode !== "copy-all" && mode !== "copy-results-only") {
    throw unknownOdsElement(source!, "/table:table", "table:mode", "invalid table source mode");
  }
  const actuate = attributeString(source, "xlink:actuate");
  if (actuate !== undefined && actuate !== "onRequest") {
    throw unknownOdsElement(
      source!,
      "/table:table",
      "xlink:actuate",
      "invalid table source actuation",
    );
  }
  if (source && attributeString(source, "xlink:type") !== "simple") {
    throw unknownOdsElement(
      source,
      "/table:table",
      "xlink:type",
      "table source link type must be simple",
    );
  }
  return {
    ...(childNamed(table, "table:title")
      ? { tableTitle: textOf(childNamed(table, "table:title")) }
      : {}),
    ...(childNamed(table, "table:desc")
      ? { tableDescription: textOf(childNamed(table, "table:desc")) }
      : {}),
    ...(scenario
      ? {
          scenario: {
            ranges: requiredAttribute(scenario, "table:scenario-ranges"),
            active: requiredBoolean(scenario, "table:is-active"),
            ...optionalBoolean(scenario, "table:display-border", "displayBorder"),
            ...(attributeString(scenario, "table:border-color")
              ? { borderColor: attributeString(scenario, "table:border-color") }
              : {}),
            ...optionalBoolean(scenario, "table:copy-back", "copyBack"),
            ...optionalBoolean(scenario, "table:copy-styles", "copyStyles"),
            ...optionalBoolean(scenario, "table:copy-formulas", "copyFormulas"),
            ...(attributeString(scenario, "table:comment")
              ? { comment: attributeString(scenario, "table:comment") }
              : {}),
            ...optionalBoolean(scenario, "table:protected", "protected"),
          },
        }
      : {}),
    ...(source
      ? {
          tableSource: {
            href: requiredAttribute(source, "xlink:href"),
            ...(mode !== undefined ? { mode } : {}),
            ...(attributeString(source, "table:table-name")
              ? { tableName: attributeString(source, "table:table-name") }
              : {}),
            ...(actuate !== undefined ? { actuate } : {}),
            ...(attributeString(source, "table:filter-name")
              ? { filterName: attributeString(source, "table:filter-name") }
              : {}),
            ...(attributeString(source, "table:filter-options")
              ? { filterOptions: attributeString(source, "table:filter-options") }
              : {}),
            ...(attributeString(source, "table:refresh-delay")
              ? { refreshDelay: attributeString(source, "table:refresh-delay") }
              : {}),
          },
        }
      : {}),
  };
}

function requiredAttribute(element: Element, name: string): string {
  const value = attributeString(element, name);
  if (value === undefined)
    throw unknownOdsElement(element, "/table:table", name, "required attribute is missing");
  return value;
}

function requiredBoolean(element: Element, name: string): boolean {
  const value = optionalBooleanValue(element, name);
  if (value === undefined)
    throw unknownOdsElement(element, "/table:table", name, "required boolean is missing");
  return value;
}

function optionalBoolean(
  element: Element,
  attribute: string,
  property: "displayBorder" | "copyBack" | "copyStyles" | "copyFormulas" | "protected",
): Record<string, boolean> {
  const value = optionalBooleanValue(element, attribute);
  return value === undefined ? {} : { [property]: value };
}

function optionalBooleanValue(element: Element, name: string): boolean | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  if (value === "true" || value === "1") return true;
  if (value === "false" || value === "0") return false;
  throw unknownOdsElement(element, "/table:table", name, "invalid boolean value");
}

function gnumericSelection(table: Element): SelectionOptions[] | undefined {
  const container = childNamed(table, "gnm:selections");
  if (!container) return undefined;
  const cursorColumn = attributeNumber(container, "gnm:cursor-col") ?? 0;
  const cursorRow = attributeNumber(container, "gnm:cursor-row") ?? 0;
  const ranges = childrenNamed(container, "gnm:selection").map((selection) => {
    const startColumn = (attributeNumber(selection, "gnm:start-col") ?? cursorColumn) + 1;
    const startRow = (attributeNumber(selection, "gnm:start-row") ?? cursorRow) + 1;
    const endColumn = (attributeNumber(selection, "gnm:end-col") ?? startColumn - 1) + 1;
    const endRow = (attributeNumber(selection, "gnm:end-row") ?? startRow - 1) + 1;
    return cellAddress(startColumn, startRow) === cellAddress(endColumn, endRow)
      ? cellAddress(startColumn, startRow)
      : `${cellAddress(startColumn, startRow)}:${cellAddress(endColumn, endRow)}`;
  });
  return [
    {
      activeCell: cellAddress(cursorColumn + 1, cursorRow + 1),
      ...(ranges.length ? { sqref: ranges.join(" ") } : {}),
    },
  ];
}

type ParsedSelectionRange = {
  startColumn: number;
  startRow: number;
  endColumn: number;
  endRow: number;
};

function gnumericSelectionXml(selections: readonly SelectionOptions[]): string[] {
  const selection = selections[0];
  if (!selection) return [];
  const active = parseCellReference(selection.activeCell);
  const ranges: (ParsedSelectionRange | undefined)[] = (
    selection.sqref ??
    selection.activeCell ??
    ""
  )
    .split(/\s+/u)
    .filter(Boolean)
    .map((range) => {
      const [start, end = start] = range.split(":") as [string, string?];
      const startCell = parseCellReference(start);
      const endCell = parseCellReference(end);
      if (!startCell || !endCell) return undefined;
      return {
        startColumn: startCell.column,
        startRow: startCell.row,
        endColumn: endCell.column,
        endRow: endCell.row,
      };
    });
  const firstRange = ranges[0];
  const cursor =
    active ??
    (firstRange
      ? {
          column: firstRange.startColumn,
          row: firstRange.startRow,
        }
      : undefined);
  if (!cursor) return [];
  return [
    xmlElement(
      "gnm:selections",
      {
        "gnm:cursor-col": cursor.column - 1,
        "gnm:cursor-row": cursor.row - 1,
      },
      ranges
        .filter((range): range is ParsedSelectionRange => range !== undefined)
        .map((range) =>
          xmlElement("gnm:selection", {
            "gnm:start-col": range.startColumn - 1,
            "gnm:start-row": range.startRow - 1,
            "gnm:end-col": range.endColumn - 1,
            "gnm:end-row": range.endRow - 1,
          }),
        ),
    ),
  ];
}

function parseCellReference(
  reference: string | undefined,
): { column: number; row: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/u.exec(reference ?? "");
  if (!match) return undefined;
  return { column: columnNumber(match[1]!), row: Number(match[2]!) };
}

function parseCoveredCellFrame(frame: Element): CellOptions["graphics"] {
  const object = childNamed(frame, "draw:object") ?? childNamed(frame, "draw:object-ole");
  const image = childNamed(frame, "draw:image");
  const href = attributeString(object ?? image, "xlink:href");
  if (!href) return undefined;
  assertElementNames(
    frame,
    [
      "draw:name",
      "svg:title",
      "svg:x",
      "svg:y",
      "svg:width",
      "svg:height",
      "draw:z-index",
      "table:end-cell-address",
      "draw:id",
      "table:end-x",
      "table:end-y",
      "draw:style-name",
      "draw:text-style-name",
      "draw:transform",
      "drawooo:display",
      "loext:p",
      "xmlns:presentation",
      "style:rel-width",
      "style:rel-height",
    ],
    ["loext:p", "draw:image", "draw:object", "draw:object-ole", "svg:title", "svg:desc"],
  );
  const common = {
    href: object
      ? href.replace(/^\.\//, "")
      : isExternalOdfUrl(href)
        ? href
        : href.replace(/^\//, ""),
    x: lengthToEmu(attributeString(frame, "svg:x")),
    y: lengthToEmu(attributeString(frame, "svg:y")),
    width: lengthToEmu(attributeString(frame, "svg:width")),
    height: lengthToEmu(attributeString(frame, "svg:height")),
    zIndex: attributeNumber(frame, "draw:z-index"),
    endCellAddress: attributeString(frame, "table:end-cell-address"),
    endX: lengthToEmu(attributeString(frame, "table:end-x")),
    endY: lengthToEmu(attributeString(frame, "table:end-y")),
    transform: attributeString(frame, "draw:transform"),
    display: attributeString(frame, "drawooo:display") as "printer" | undefined,
  };
  const metadata = {
    name: attributeString(frame, "draw:name"),
    title: textOf(childNamed(frame, "svg:title")),
    description: textOf(childNamed(frame, "svg:desc")),
    id: attributeString(frame, "draw:id"),
    styleName: attributeString(frame, "draw:style-name"),
    textStyleName: attributeString(frame, "draw:text-style-name"),
    relativeWidth: attributeString(frame, "style:rel-width"),
    relativeHeight: attributeString(frame, "style:rel-height"),
  };
  const graphic = object
    ? { type: "object" as const, ...common, ...metadata }
    : { type: "image" as const, ...common, data: "", ...metadata };
  return [graphic];
}

export function validateTableChildren(table: Element): void {
  for (const child of table.elements ?? []) {
    if (child.type !== "element" || !child.name) continue;
    if (
      child.name === "table:table-column" ||
      child.name === "table:table-row" ||
      child.name === "table:title" ||
      child.name === "table:desc" ||
      child.name === "table:table-source" ||
      child.name === "table:scenario" ||
      child.name === "office:forms" ||
      child.name === "table:shapes" ||
      child.name === "table:named-expressions" ||
      child.name === "loext:table-protection" ||
      child.name === "calcext:conditional-formats" ||
      child.name === "calcext:sparkline-groups" ||
      child.name === "table:detective" ||
      child.name === "gnm:selections" ||
      child.name === "draw:frame" ||
      COLUMN_CONTAINERS.has(child.name) ||
      ROW_CONTAINERS.has(child.name)
    ) {
      continue;
    }
    throw unknownOdsElement(
      child,
      `/office:document-content/office:body/office:spreadsheet/table:table[@table:name="${attributeString(table, "table:name") ?? ""}"]`,
      child.name,
      "element has no canonical WorksheetOptions mapping",
    );
  }
}

export function tableColumns(table: Element): Element[] {
  return (table.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-column") return [child];
    if (COLUMN_CONTAINERS.has(child.name ?? "")) return tableColumns(child);
    return [];
  });
}

export function tableRows(table: Element): Element[] {
  return (table.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-row") return [child];
    if (ROW_CONTAINERS.has(child.name ?? "")) return tableRows(child);
    return [];
  });
}

export function parseColumn(
  element: Element,
  column: number,
  dimensions: Map<string, DimensionStyle>,
): ColumnOptions {
  const style = dimensions.get(attributeString(element, "table:style-name") ?? "");
  return {
    min: column,
    max: column,
    width: style?.width,
    bestFit: style?.bestFit,
    hidden: tableVisibility(element, "table:visibility") !== "visible" || style?.hidden,
  };
}

function parseRows(
  table: Element,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  graphicStyles: Map<string, GraphicStyle>,
  comments: CommentOptions[],
  mergeCells: MergeCellOptions[],
  shapes: WorksheetOptions["shapes"],
  textStyles: Map<string, RichTextRunPropertiesOptions>,
  hyperlinks: HyperlinkOptions[],
  groups: WorksheetOptions["groups"],
): { rows: OdsRowOptions[]; defaultRowHeight?: number } {
  const rows: OdsRowOptions[] = [];
  let defaultRowHeight: number | undefined;
  let rowNumber = 1;
  for (const row of tableRows(table)) {
    const repeat = repeatCount(row, "table:number-rows-repeated");
    const mergeStart = mergeCells.length;
    const parsed = parseRow(
      row,
      rowNumber,
      dimensions,
      cellStyles,
      graphicStyles,
      comments,
      mergeCells,
      shapes,
      textStyles,
      hyperlinks,
      groups,
    );
    const hasRowData = (parsed.cells ?? []).some(hasCellContent) || parsed.hidden === true;
    if (repeat > MAX_ROWS && hasRowData) {
      if (mergeCells.length !== mergeStart) {
        throw new OdsParseError(
          "Repeated rows with merged cells exceed the canonical worksheet limits",
          "content.xml",
          "office:body/office:spreadsheet/table:table",
          "table:number-rows-repeated",
          "unsupported-repeated-merge",
        );
      }
      rows.push({ ...parsed, repeat });
      rowNumber += repeat;
      continue;
    }
    if (isLargeEmptyRepeat(repeat, hasRowData)) {
      if (typeof parsed.height === "number" && defaultRowHeight === undefined) {
        defaultRowHeight = parsed.height;
      }
      rowNumber += repeat;
      continue;
    }
    if (repeat > 1) {
      const sourceMergeCells = mergeCells.slice(mergeStart);
      mergeCells.splice(mergeStart, sourceMergeCells.length);
      const repeatRowSpan = Math.max(
        1,
        ...sourceMergeCells.map((merge) => rangeRowSpan(merge.ref)),
      );
      for (let index = 0; index < repeat; index += 1) {
        const rowOffset = index * repeatRowSpan;
        const currentRow = rowNumber + rowOffset;
        const rowMergeCells = sourceMergeCells.map((mergeCell) => ({
          ...mergeCell,
          ref: shiftRangeRowReference(mergeCell.ref, rowOffset),
        }));
        mergeCells.push(...rowMergeCells);
        rows.push({
          ...parsed,
          rowNumber: currentRow,
          cells: (parsed.cells ?? []).map((cell) => renumberCell(cell, currentRow)),
        });
      }
      rowNumber += repeat * repeatRowSpan;
      continue;
    }
    rows.push(parsed);
    rowNumber += repeat;
  }
  return {
    rows: rows.map((row) => {
      if (row.height !== defaultRowHeight) return row;
      const { height: _height, ...normalized } = row;
      return normalized;
    }),
    defaultRowHeight,
  };
}

function renumberCell(cell: CellOptions, rowNumber: number): CellOptions {
  const next = {
    ...cell,
    reference: replaceRowReference(cell.reference ?? "", rowNumber),
    ...(cell.covered
      ? {
          covered: {
            ...cell.covered,
            reference: replaceRowReference(cell.covered.reference, rowNumber),
          },
        }
      : {}),
  };
  return next;
}

function replaceRowReference(reference: string, rowNumber: number): string {
  const match = /^([A-Z]+)(\d+)$/u.exec(reference);
  return match ? `${match[1]}${rowNumber}` : reference;
}

function shiftRangeRowReference(reference: string, rowOffset: number): string {
  return reference.replace(
    /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/u,
    (_, startColumn: string, startRow: string, endColumn: string, endRow: string) =>
      `${startColumn}${Number(startRow) + rowOffset}:${endColumn}${Number(endRow) + rowOffset}`,
  );
}

function rangeRowSpan(reference: string): number {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/u.exec(reference);
  if (!match) return 1;
  return Math.max(1, Number(match[4]) - Number(match[2]) + 1);
}

export function parseRow(
  row: Element,
  rowNumber: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  graphicStyles: Map<string, GraphicStyle>,
  comments: CommentOptions[],
  mergeCells: MergeCellOptions[] = [],
  shapes: WorksheetOptions["shapes"] = [],
  textStyles: Map<string, RichTextRunPropertiesOptions> = new Map(),
  hyperlinks: HyperlinkOptions[] = [],
  groups: WorksheetOptions["groups"] = [],
): OdsRowOptions {
  const style = dimensions.get(attributeString(row, "table:style-name") ?? "");
  const cellStartColumns: number[] = [];
  let columnNumber = 1;
  for (const cell of row.elements ?? []) {
    if (cell.type !== "element" || !cell.name) continue;
    const repeat = repeatCount(cell, "table:number-columns-repeated");
    cellStartColumns.push(columnNumber);
    columnNumber += repeat;
  }
  return {
    rowNumber,
    height: style?.height,
    hidden: tableVisibility(row, "table:visibility") !== "visible" || style?.hidden,
    cells: (row.elements ?? [])
      .flatMap((cell, cellIndex) => {
        const repeat = repeatCount(cell, "table:number-columns-repeated");
        const startColumn = cellStartColumns[cellIndex]!;
        if (cell.type !== "element" || !cell.name) return undefined;
        if (cell.name === "table:covered-table-cell") {
          assertElementNames(
            cell,
            [
              "table:style-name",
              "office:value-type",
              "calcext:value-type",
              "table:number-columns-repeated",
              "table:content-validation-name",
              "table:formula",
              "table:number-matrix-columns-spanned",
              "table:number-matrix-rows-spanned",
              "office:value",
              "office:string-value",
              "office:currency",
              "office:date-value",
              "office:boolean-value",
              "office:time-value",
            ],
            ["text:p", "draw:frame"],
          );
          const coveredStyle = cellStyles.get(attributeString(cell, "table:style-name") ?? "");
          const coveredCell: OdsCellOptions = {
            reference: `${columnName(startColumn)}${rowNumber}`,
            covered: {
              reference: `${columnName(startColumn)}${rowNumber}`,
              ...(childNamed(cell, "text:p") ? { text: textOf(childNamed(cell, "text:p")) } : {}),
              ...(attributeString(cell, "office:value-type")
                ? {
                    valueType: attributeString(cell, "office:value-type") as
                      | "float"
                      | "string"
                      | "date"
                      | "boolean"
                      | "percentage"
                      | "currency"
                      | "time",
                  }
                : {}),
            },
            ...(coveredStyle ? { style: coveredStyle } : {}),
          };
          const graphics = childNamed(cell, "draw:frame")
            ? parseCoveredCellFrame(childNamed(cell, "draw:frame")!)
            : undefined;
          if (graphics) coveredCell.graphics = graphics;
          if (repeat > MAX_COLUMNS) {
            return [{ ...coveredCell, repeat }];
          }
          return Array.from({ length: repeat }, (_, repeatedIndex) => ({
            ...coveredCell,
            reference: `${columnName(startColumn + repeatedIndex)}${rowNumber}`,
            covered: {
              ...coveredCell.covered,
              reference: `${columnName(startColumn + repeatedIndex)}${rowNumber}`,
            },
            ...(graphics ? { graphics: structuredClone(graphics) } : {}),
          }));
        }
        if (cell.name !== "table:table-cell") {
          throw unknownOdsElement(
            cell,
            `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${rowNumber}]`,
            cell.name,
            "element has no canonical CellOptions mapping",
          );
        }
        const firstCell: OdsCellOptions = parseCell(
          cell,
          rowNumber,
          startColumn,
          cellStyles,
          [],
          textStyles,
        );
        const hasMergeSpan =
          (attributeNumber(cell, "table:number-columns-spanned") ?? 1) > 1 ||
          (attributeNumber(cell, "table:number-rows-spanned") ?? 1) > 1;
        const hasCellSurface =
          hasCellContent(firstCell) ||
          hasMergeSpan ||
          hasVisibleCellStyle(typeof firstCell.style === "number" ? undefined : firstCell.style);
        if (repeat > MAX_COLUMNS && hasCellSurface) {
          if (
            hasMergeSpan ||
            (cell.elements ?? []).some(
              (child) =>
                child.name === "office:annotation" ||
                child.name === "draw:frame" ||
                child.name === "draw:g",
            ) ||
            parseCellHyperlinks(cell, firstCell.reference ?? "").length > 0
          ) {
            throw new OdsParseError(
              "Repeated cells with per-cell metadata exceed the canonical worksheet limits",
              "content.xml",
              "office:body/office:spreadsheet/table:table",
              "table:number-columns-repeated",
              "unsupported-repeated-cell-metadata",
            );
          }
          return [{ ...firstCell, repeat }];
        }
        if (isLargeEmptyRepeat(repeat, hasCellSurface)) return [];
        return Array.from({ length: repeat }, (_, repeatedIndex) => {
          const column = startColumn + repeatedIndex;
          const parsed = parseCell(cell, rowNumber, column, cellStyles, mergeCells, textStyles);
          hyperlinks.push(...parseCellHyperlinks(cell, parsed.reference ?? ""));
          for (const child of cell.elements ?? []) {
            if (child.name === "office:annotation") {
              comments.push({
                cell: parsed.reference ?? "",
                author: attributeString(child, "dc:creator") ?? "",
                text: childrenNamed(child, "text:p")
                  .map((paragraph) => textOf(paragraph))
                  .join("\n"),
              });
            }
            if (child.name === "draw:frame") {
              const object =
                childNamed(child, "draw:object") ?? childNamed(child, "draw:object-ole");
              const image = childNamed(child, "draw:image");
              const href = attributeString(object ?? image, "xlink:href");
              const description = textOf(childNamed(child, "svg:desc"));
              const title = textOf(childNamed(child, "svg:title"));
              const updateNotification = childNamed(child, "loext:p");
              const notifyOnUpdateOfRanges = attributeString(
                updateNotification,
                "draw:notify-on-update-of-ranges",
              );
              if (updateNotification)
                assertElementNames(updateNotification, ["draw:notify-on-update-of-ranges"]);
              if (!href) {
                throw unknownOdsElement(
                  child,
                  `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${rowNumber}]/table:table-cell[${column}]`,
                  child.name!,
                  "graphic href is required",
                );
              }
              assertElementNames(
                child,
                [
                  "draw:name",
                  "svg:title",
                  "svg:x",
                  "svg:y",
                  "svg:width",
                  "svg:height",
                  "draw:z-index",
                  "table:end-cell-address",
                  "draw:id",
                  "table:end-x",
                  "table:end-y",
                  "draw:style-name",
                  "draw:text-style-name",
                  "draw:transform",
                  "drawooo:display",
                  "loext:p",
                  "xmlns:presentation",
                  "style:rel-width",
                  "style:rel-height",
                ],
                [
                  "loext:p",
                  "draw:image",
                  "draw:object",
                  "draw:object-ole",
                  "svg:title",
                  "svg:desc",
                ],
              );
              parsed.graphics = [
                ...(parsed.graphics ?? []),
                object
                  ? {
                      type: "object",
                      href: href.replace(/^\.\//, "").replace(/^\//, ""),
                      ...(attributeString(child, "draw:name")
                        ? { name: attributeString(child, "draw:name") }
                        : {}),
                      ...(description ? { description } : {}),
                      ...(title ? { title } : {}),
                      transform: attributeString(child, "draw:transform"),
                      display: graphicDisplay(child),
                      ...(notifyOnUpdateOfRanges === undefined ? {} : { notifyOnUpdateOfRanges }),
                      x: lengthToEmu(attributeString(child, "svg:x")),
                      y: lengthToEmu(attributeString(child, "svg:y")),
                      width: lengthToEmu(attributeString(child, "svg:width")),
                      height: lengthToEmu(attributeString(child, "svg:height")),
                      ...(attributeNumber(child, "draw:z-index") !== undefined
                        ? { zIndex: attributeNumber(child, "draw:z-index") }
                        : {}),
                      ...(attributeString(child, "table:end-cell-address")
                        ? { endCellAddress: attributeString(child, "table:end-cell-address") }
                        : {}),
                      ...(attributeString(child, "draw:id")
                        ? { id: attributeString(child, "draw:id") }
                        : {}),
                      ...(attributeString(child, "table:end-x") !== undefined
                        ? { endX: lengthToEmu(attributeString(child, "table:end-x")) }
                        : {}),
                      ...(attributeString(child, "table:end-y") !== undefined
                        ? { endY: lengthToEmu(attributeString(child, "table:end-y")) }
                        : {}),
                      ...(attributeString(child, "draw:style-name")
                        ? { styleName: attributeString(child, "draw:style-name") }
                        : {}),
                      ...(attributeString(child, "draw:text-style-name")
                        ? { textStyleName: attributeString(child, "draw:text-style-name") }
                        : {}),
                      ...(attributeString(child, "style:rel-width")
                        ? { relativeWidth: attributeString(child, "style:rel-width") }
                        : {}),
                      ...(attributeString(child, "style:rel-height")
                        ? { relativeHeight: attributeString(child, "style:rel-height") }
                        : {}),
                    }
                  : {
                      type: "image",
                      href: isExternalOdfUrl(href) ? href : href.replace(/^\//, ""),
                      data: "",
                      ...(attributeString(child, "draw:name")
                        ? { name: attributeString(child, "draw:name") }
                        : {}),
                      ...(description ? { description } : {}),
                      ...(title ? { title } : {}),
                      transform: attributeString(child, "draw:transform"),
                      display: graphicDisplay(child),
                      ...(notifyOnUpdateOfRanges === undefined ? {} : { notifyOnUpdateOfRanges }),
                      x: lengthToEmu(attributeString(child, "svg:x")),
                      y: lengthToEmu(attributeString(child, "svg:y")),
                      width: lengthToEmu(attributeString(child, "svg:width")),
                      height: lengthToEmu(attributeString(child, "svg:height")),
                      ...(attributeNumber(child, "draw:z-index") !== undefined
                        ? { zIndex: attributeNumber(child, "draw:z-index") }
                        : {}),
                      ...(attributeString(child, "table:end-cell-address")
                        ? { endCellAddress: attributeString(child, "table:end-cell-address") }
                        : {}),
                      ...(attributeString(child, "draw:id")
                        ? { id: attributeString(child, "draw:id") }
                        : {}),
                      ...(attributeString(child, "table:end-x") !== undefined
                        ? { endX: lengthToEmu(attributeString(child, "table:end-x")) }
                        : {}),
                      ...(attributeString(child, "table:end-y") !== undefined
                        ? { endY: lengthToEmu(attributeString(child, "table:end-y")) }
                        : {}),
                      ...(attributeString(child, "draw:style-name")
                        ? { styleName: attributeString(child, "draw:style-name") }
                        : {}),
                      ...(attributeString(child, "draw:text-style-name")
                        ? { textStyleName: attributeString(child, "draw:text-style-name") }
                        : {}),
                      ...(attributeString(child, "style:rel-width")
                        ? { relativeWidth: attributeString(child, "style:rel-width") }
                        : {}),
                      ...(attributeString(child, "style:rel-height")
                        ? { relativeHeight: attributeString(child, "style:rel-height") }
                        : {}),
                    },
              ];
            }
            if (
              child.name === "draw:rect" ||
              child.name === "draw:control" ||
              child.name === "draw:ellipse" ||
              child.name === "draw:circle" ||
              child.name === "draw:line" ||
              child.name === "draw:measure" ||
              child.name === "draw:polygon" ||
              child.name === "draw:polyline" ||
              child.name === "draw:custom-shape"
            ) {
              shapes?.push(parseDrawingShape(child, rowNumber, column, graphicStyles, textStyles));
            }
            if (child.name === "draw:g") {
              groups?.push(parseDrawingGroup(child, rowNumber, column, graphicStyles, textStyles));
            }
          }
          return parsed;
        });
      })
      .filter((cell): cell is CellOptions => cell !== undefined),
  };
}

function tableVisibility(element: Element, name: string): "visible" | "collapse" | "filter" {
  const value = attributeString(element, name);
  if (value === undefined || value === "visible" || value === "collapse" || value === "filter")
    return value ?? "visible";
  throw unknownOdsElement(
    element,
    "/office:document-content/office:body/office:spreadsheet",
    name,
    `invalid visibility value: ${value}`,
  );
}

export function parseCell(
  cell: Element,
  row: number,
  column: number,
  cellStyles: Map<string, StyleOptions>,
  mergeCells: MergeCellOptions[] = [],
  textStyles: Map<string, RichTextRunPropertiesOptions> = new Map(),
): CellOptions {
  const result: CellOptions = { reference: `${columnName(column)}${row}` };
  const columnSpan = attributeNumber(cell, "table:number-columns-spanned") ?? 1;
  const rowSpan = attributeNumber(cell, "table:number-rows-spanned") ?? 1;
  if (
    !Number.isInteger(columnSpan) ||
    !Number.isInteger(rowSpan) ||
    columnSpan < 1 ||
    rowSpan < 1 ||
    columnSpan - 1 > Number.MAX_SAFE_INTEGER - column ||
    rowSpan - 1 > Number.MAX_SAFE_INTEGER - row
  ) {
    throw new OdsParseError(
      "Invalid OpenDocument spreadsheet cell span",
      "content.xml",
      `office:body/office:spreadsheet/table:table/table:table-row[${row}]`,
      cell.name ?? "",
      "invalid-cell-span",
    );
  }
  if (columnSpan > 1 || rowSpan > 1) {
    const reference = `${columnName(column)}${row}:${columnName(column + columnSpan - 1)}${row + rowSpan - 1}`;
    if (!mergeCells.some((mergeCell) => mergeCell.ref === reference)) {
      mergeCells.push({ ref: reference });
    }
  }
  for (const child of cell.elements ?? []) {
    if (
      child.type === "element" &&
      ![
        "text:p",
        "loext:p",
        "office:annotation",
        "table:cell-range-source",
        "table:detective",
        "draw:frame",
        "draw:rect",
        "draw:control",
        "draw:ellipse",
        "draw:circle",
        "draw:line",
        "draw:measure",
        "draw:polygon",
        "draw:polyline",
        "draw:custom-shape",
        "draw:g",
      ].includes(child.name ?? "")
    ) {
      throw unknownOdsElement(
        child,
        `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
        child.name ?? "",
        "element has no canonical CellOptions mapping",
      );
    }
  }
  const source = childNamed(cell, "table:cell-range-source");
  if (source) result.cellRangeSource = parseCellRangeSource(source);
  const detective = childNamed(cell, "table:detective");
  if (detective) result.detective = parseTableDetective(detective);
  const style = cellStyles.get(attributeString(cell, "table:style-name") ?? "");
  if (style) result.style = style;
  const formula = attributeString(cell, "table:formula");
  if (formula) {
    result.formula = formula.replace(/^(?:(?:of|ooow|msoxl):)?=/u, "");
    const columnsSpanned = attributeNumber(cell, "table:number-matrix-columns-spanned");
    const rowsSpanned = attributeNumber(cell, "table:number-matrix-rows-spanned");
    if ((columnsSpanned ?? 1) > 1 || (rowsSpanned ?? 1) > 1) {
      result.formula = {
        formula: result.formula,
        type: "array",
        reference: `${result.reference}:${columnName(column + (columnsSpanned ?? 1) - 1)}${
          row + (rowsSpanned ?? 1) - 1
        }`,
      };
    }
  }
  const valueType = attributeString(cell, "office:value-type");
  if (valueType === "float") {
    const value = attributeNumber(cell, "office:value");
    if (value === undefined) throw invalidCellValue(cell, row, column, "office:value");
    result.value = value;
  } else if (valueType === "percentage") {
    const value = attributeNumber(cell, "office:value");
    if (value === undefined) throw invalidCellValue(cell, row, column, "office:value");
    result.value = value;
    result.style = styleWithNumFmt(result.style, "0%");
  } else if (valueType === "currency") {
    const value = attributeNumber(cell, "office:value");
    if (value === undefined) throw invalidCellValue(cell, row, column, "office:value");
    result.value = value;
    const currency = attributeString(cell, "office:currency");
    if (currency) result.style = styleWithNumFmt(result.style, `[$${currency}] #,##0.00`);
  } else if (valueType === "time") {
    const value = durationToSerial(attributeString(cell, "office:time-value"));
    if (value === undefined)
      throw unknownOdsElement(
        cell,
        `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
        "office:time-value",
        "invalid time duration",
      );
    result.value = value;
    result.style = styleWithNumFmt(result.style, "[hh]:mm:ss");
  } else if (valueType === "boolean") {
    const value = attributeString(cell, "office:boolean-value");
    if (value !== "true" && value !== "false" && value !== "1" && value !== "0")
      throw unknownOdsElement(
        cell,
        `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
        "office:boolean-value",
        "invalid boolean value",
      );
    result.value = value === "true" || value === "1";
  } else if (valueType === "date") {
    const source = attributeString(cell, "office:date-value");
    const normalized =
      source === undefined || /(?:Z|[+-]\d{2}:\d{2})$/u.test(source) ? source : `${source}Z`;
    const date = normalized === undefined ? undefined : new Date(normalized);
    if (!date || Number.isNaN(date.getTime()))
      throw unknownOdsElement(
        cell,
        `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
        "office:date-value",
        "invalid date value",
      );
    result.value = date;
  } else if (valueType === "string") {
    result.value = parseCellText(cell, textStyles);
  } else if (valueType === "error") {
    const error =
      attributeString(cell, "office:string-value") ?? attributeString(cell, "office:value");
    if (error === undefined) throw invalidCellValue(cell, row, column, "office:string-value");
    result.error = error;
  } else if (valueType !== undefined) {
    throw unknownOdsElement(
      cell,
      `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
      "office:value-type",
      `unsupported value type: ${valueType}`,
    );
  }
  return result;
}

function parseCellText(
  cell: Element,
  textStyles: Map<string, RichTextRunPropertiesOptions>,
): string | RichTextOptions {
  const paragraphs = childrenNamed(cell, "text:p");
  const hasStyledText = paragraphs.some((paragraph) =>
    (paragraph.elements ?? []).some(
      (child) => child.type === "element" && child.name === "text:span",
    ),
  );
  const hasStructuredText = paragraphs.some((paragraph) =>
    (paragraph.elements ?? []).some(
      (child) =>
        child.type === "element" &&
        ["text:span", "text:s", "text:tab", "text:line-break", "text:a"].includes(child.name ?? ""),
    ),
  );
  if (!hasStructuredText) return paragraphs.map(textOf).join("\n");
  const runs: RichTextRunOptions[] = [];
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    if (paragraphIndex > 0) runs.push({ text: "\n" });
    appendTextRuns(paragraph, textStyles, runs);
  }
  if (!hasStyledText) return runs.map((run) => run.text).join("");
  return { runs };
}

function appendTextRuns(
  element: Element,
  textStyles: Map<string, RichTextRunPropertiesOptions>,
  runs: RichTextRunOptions[],
  inherited?: RichTextRunPropertiesOptions,
): void {
  for (const child of element.elements ?? []) {
    if (child.type === "text") {
      const text = String(child.text ?? "");
      if (text.length > 0) pushTextRun(runs, text, inherited);
      continue;
    }
    if (child.name === "text:s") {
      const count = attributeNumber(child, "text:c") ?? 1;
      if (!Number.isInteger(count) || count < 1)
        throw unknownOdsElement(child, "/table:table-cell/text:p", "text:c", "invalid space count");
      pushTextRun(runs, " ".repeat(count), inherited);
      continue;
    }
    if (child.name === "text:tab") {
      pushTextRun(runs, "\t", inherited);
      continue;
    }
    if (child.name === "text:line-break") {
      pushTextRun(runs, "\n", inherited);
      continue;
    }
    if (child.name === "text:span") {
      const styleName = attributeString(child, "text:style-name");
      const properties = styleName ? { ...inherited, ...textStyles.get(styleName) } : inherited;
      appendTextRuns(child, textStyles, runs, properties);
      continue;
    }
    if (child.name === "text:a") appendTextRuns(child, textStyles, runs, inherited);
  }
}

function parseCellHyperlinks(cell: Element, reference: string): HyperlinkOptions[] {
  return childrenNamed(cell, "text:p").flatMap((paragraph) =>
    (paragraph.elements ?? []).flatMap((child) => {
      if (child.type !== "element" || child.name !== "text:a") return [];
      const href = attributeString(child, "xlink:href");
      if (href === undefined) return [];
      return [
        {
          cell: reference,
          ...(href.startsWith("#") ? { location: href.slice(1) } : { url: href }),
          display: textOf(child),
        },
      ];
    }),
  );
}

function pushTextRun(
  runs: RichTextRunOptions[],
  text: string,
  properties: RichTextRunPropertiesOptions | undefined,
): void {
  const previous = runs[runs.length - 1];
  if (previous && JSON.stringify(previous.properties) === JSON.stringify(properties)) {
    previous.text += text;
    return;
  }
  runs.push({ text, ...(properties ? { properties } : {}) });
}

function invalidCellValue(cell: Element, row: number, column: number, name: string): OdsParseError {
  return unknownOdsElement(
    cell,
    `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
    name,
    "required cell value is missing or invalid",
  );
}

function parseCellRangeSource(source: Element): CellRangeSourceOptions {
  assertElementNames(source, [
    "table:name",
    "table:last-column-spanned",
    "table:last-row-spanned",
    "xlink:type",
    "xlink:href",
    "xlink:actuate",
    "table:filter-name",
    "table:filter-options",
    "table:refresh-delay",
  ]);
  const name = attributeString(source, "table:name");
  const lastColumnSpanned = attributeNumber(source, "table:last-column-spanned");
  const lastRowSpanned = attributeNumber(source, "table:last-row-spanned");
  const href = attributeString(source, "xlink:href");
  if (
    !name ||
    !lastColumnSpanned ||
    !lastRowSpanned ||
    href === undefined ||
    attributeString(source, "xlink:type") !== "simple"
  ) {
    throw unknownOdsElement(
      source,
      "office:body/office:spreadsheet/table:table-cell",
      source.name ?? "",
      "invalid linked cell range",
    );
  }
  return {
    name,
    lastColumnSpanned,
    lastRowSpanned,
    href,
    ...(attributeString(source, "xlink:actuate") === "onRequest"
      ? { actuate: "onRequest" as const }
      : {}),
    ...(attributeString(source, "table:filter-name")
      ? { filterName: attributeString(source, "table:filter-name") }
      : {}),
    ...(attributeString(source, "table:filter-options")
      ? { filterOptions: attributeString(source, "table:filter-options") }
      : {}),
    ...(attributeString(source, "table:refresh-delay")
      ? { refreshDelay: attributeString(source, "table:refresh-delay") }
      : {}),
  };
}

const HIGHLIGHT_DIRECTION_VALUES: Record<string, DetectiveHighlightedRangeDirection> = {
  "from-another-table": "fromAnotherTable",
  "to-another-table": "toAnotherTable",
  "from-same-table": "fromSameTable",
};

const OPERATION_NAME_VALUES: Record<string, DetectiveOperationOptions["name"]> = {
  "trace-dependents": "traceDependents",
  "remove-dependents": "removeDependents",
  "trace-precedents": "tracePrecedents",
  "remove-precedents": "removePrecedents",
  "trace-errors": "traceErrors",
};

function parseTableDetective(detective: Element): TableDetectiveOptions {
  assertElementNames(detective, [], ["table:highlighted-range", "table:operation"]);
  const highlightedRanges = childrenNamed(detective, "table:highlighted-range").map(
    (range): DetectiveHighlightedRangeOptions => {
      assertElementNames(range, [
        "table:cell-range-address",
        "table:direction",
        "table:contains-error",
        "table:marked-invalid",
      ]);
      const direction = HIGHLIGHT_DIRECTION_VALUES[attributeString(range, "table:direction") ?? ""];
      const markedInvalid = attributeString(range, "table:marked-invalid");
      if (markedInvalid !== undefined) {
        if (
          direction ||
          markedInvalid === undefined ||
          !["true", "false"].includes(markedInvalid)
        ) {
          throw unknownOdsElement(
            range,
            "office:body/office:spreadsheet/table:table-cell/table:detective",
            range.name ?? "",
            "invalid highlighted range",
          );
        }
        return { markedInvalid: markedInvalid === "true" };
      }
      if (!direction) {
        throw unknownOdsElement(
          range,
          "office:body/office:spreadsheet/table:table-cell/table:detective",
          range.name ?? "",
          "highlight direction is required",
        );
      }
      const containsError = attributeString(range, "table:contains-error");
      if (containsError !== undefined && !["true", "false"].includes(containsError)) {
        throw unknownOdsElement(
          range,
          "office:body/office:spreadsheet/table:table-cell/table:detective",
          range.name ?? "",
          "invalid contains-error value",
        );
      }
      return {
        ...(attributeString(range, "table:cell-range-address")
          ? { cellRangeAddress: attributeString(range, "table:cell-range-address") }
          : {}),
        direction,
        ...(containsError !== undefined ? { containsError: containsError === "true" } : {}),
      };
    },
  );
  const operations = childrenNamed(detective, "table:operation").map(
    (operation): DetectiveOperationOptions => {
      assertElementNames(operation, ["table:name", "table:index"]);
      const name = OPERATION_NAME_VALUES[attributeString(operation, "table:name") ?? ""];
      const index = attributeNumber(operation, "table:index");
      if (!name || index === undefined || !Number.isInteger(index) || index < 0) {
        throw unknownOdsElement(
          operation,
          "office:body/office:spreadsheet/table:table-cell/table:detective",
          operation.name ?? "",
          "invalid detective operation",
        );
      }
      return { name, index };
    },
  );
  return {
    ...(highlightedRanges.length ? { highlightedRanges } : {}),
    ...(operations.length ? { operations } : {}),
  };
}

export function columnName(column: number): string {
  return spreadsheetColumnName(column);
}
