import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  graphicFill,
  graphicOutline,
  lengthToEmu,
  parseGraphicStyles,
  odfLengthToEmu,
  pushShapeStyle,
  spreadsheetColumnName,
  spreadsheetColumnNumber,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type { GraphicStyle } from "@office-open/odf";
import type {
  ConditionalFormatOptions,
  ConditionalFormatRule,
  CfvoOptions,
  DataValidationOptions,
  DxfOptions,
  IconSetType,
  PivotDataField,
  ShapeOptions,
  SheetProtectionOptions,
  StyleOptions,
  TableOptions,
  WorksheetOptions,
  WorkbookOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { parseNumberStyles } from "./numbering";
import { assertElementNames, unknownOdsElement } from "./workbook";

const CONTENT_PATH = "/office:document-content/office:body/office:spreadsheet";

export interface SheetAddress {
  sheet?: string;
  address: string;
}

export function parseWorksheetGraphicStyles(
  automaticStyles: Element | undefined,
  stylesDocument: Element | undefined,
): Map<string, GraphicStyle> {
  return new Map([
    ...parseGraphicStyles(automaticStyles),
    ...parseGraphicStyles(childNamed(stylesDocument, "office:styles")),
  ]);
}

export function parseDrawingShape(
  element: Element,
  row: number,
  column: number,
  graphicStyles: Map<string, GraphicStyle>,
): ShapeOptions {
  assertElementNames(
    element,
    [
      "draw:name",
      "draw:style-name",
      "draw:text-style-name",
      "draw:transform",
      "draw:z-index",
      "svg:x",
      "svg:y",
      "svg:x1",
      "svg:y1",
      "svg:x2",
      "svg:y2",
      "svg:width",
      "svg:height",
      "svg:viewBox",
      "svg:points",
      "draw:points",
      "draw:start-angle",
      "draw:end-angle",
      "draw:kind",
      "table:end-cell-address",
      "table:end-x",
      "table:end-y",
      "drawooo:display",
      "draw:control",
    ],
    ["text:p", "svg:title", "svg:desc", "draw:enhanced-geometry", "office:event-listeners"],
  );
  const isLine = element.name === "draw:line" || element.name === "draw:measure";
  const x1 = isLine
    ? lengthToEmu(attributeString(element, "svg:x1"))
    : lengthToEmu(attributeString(element, "svg:x"));
  const y1 = isLine
    ? lengthToEmu(attributeString(element, "svg:y1"))
    : lengthToEmu(attributeString(element, "svg:y"));
  const x2 = isLine ? lengthToEmu(attributeString(element, "svg:x2")) : undefined;
  const y2 = isLine ? lengthToEmu(attributeString(element, "svg:y2")) : undefined;
  const width = isLine
    ? Math.abs((x2 ?? 0) - (x1 ?? 0))
    : (lengthToEmu(attributeString(element, "svg:width")) ?? 0);
  const height = isLine
    ? Math.abs((y2 ?? 0) - (y1 ?? 0))
    : (lengthToEmu(attributeString(element, "svg:height")) ?? 0);
  const style = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const text = childrenNamed(element, "text:p")
    .map((paragraph) => textOf(paragraph))
    .join("\n");
  const endAddress = parseSheetAddress(attributeString(element, "table:end-cell-address"));
  const endCell = /^([A-Z]+)([0-9]+)$/u.exec(endAddress?.address ?? "");
  return {
    col: column,
    row,
    ...(attributeString(element, "draw:control") === undefined
      ? {}
      : { controlId: attributeString(element, "draw:control") }),
    colOffset: x1,
    rowOffset: y1,
    extentCx: width,
    extentCy: height,
    ...(endCell
      ? {
          toCol: spreadsheetColumnNumber(endCell[1] ?? "A"),
          toRow: Number(endCell[2] ?? row),
          toColOffset: lengthToEmu(attributeString(element, "table:end-x")),
          toRowOffset: lengthToEmu(attributeString(element, "table:end-y")),
        }
      : {}),
    name: attributeString(element, "draw:name") ?? `${element.name}:${column}:${row}`,
    zOrder: attributeNumber(element, "draw:z-index"),
    properties: {
      x: x1,
      y: y1,
      width,
      height,
      geometry:
        element.name === "draw:rect" || element.name === "draw:control"
          ? "rect"
          : element.name === "draw:ellipse" || element.name === "draw:circle"
            ? "ellipse"
            : isLine
              ? "line"
              : null,
      fill: graphicFill(style),
      outline: graphicOutline(style),
    },
    ...(text ? { textBody: { text } } : {}),
  };
}

export function parseTableProtection(table: Element, worksheet: WorksheetOptions): void {
  const protection = childNamed(table, "loext:table-protection");
  if (!protection) return;
  assertElementNames(protection, [
    "loext:select-protected-cells",
    "loext:select-unprotected-cells",
    "loext:insert-rows",
    "loext:delete-rows",
  ]);
  const options: SheetProtectionOptions = {
    sheet: true,
    selectLockedCells: attributeString(protection, "loext:select-protected-cells") === "false",
    selectUnlockedCells: attributeString(protection, "loext:select-unprotected-cells") === "false",
    insertRows: attributeString(protection, "loext:insert-rows") === "false",
    deleteRows: attributeString(protection, "loext:delete-rows") === "false",
  };
  worksheet.protection = options;
}

export function cellShapeXml(
  shapes: WorksheetOptions["shapes"],
  reference: string,
  styles: string[],
): string[] {
  return (shapes ?? [])
    .filter((shape) => `${spreadsheetColumnName(shape.col)}${shape.row}` === reference)
    .map((shape) => {
      const geometry = shape.properties.geometry;
      const isLine = geometry === "line";
      const elementName =
        shape.controlId !== undefined
          ? "draw:control"
          : geometry === "rect"
            ? "draw:rect"
            : geometry === "ellipse"
              ? "draw:ellipse"
              : "draw:custom-shape";
      const styleName = pushShapeStyle(shape.properties.fill, shape.properties.outline, styles);
      return xmlElement(
        elementName,
        {
          "draw:name": shape.name,
          ...(shape.controlId === undefined ? {} : { "draw:control": shape.controlId }),
          "draw:style-name": styleName,
          "draw:z-index": shape.zOrder,
          ...(isLine
            ? {
                "svg:x1": emuToLength(shape.colOffset ?? 0),
                "svg:y1": emuToLength(shape.rowOffset ?? 0),
                "svg:x2": emuToLength(
                  odfLengthToEmu(shape.colOffset ?? 0) + odfLengthToEmu(shape.extentCx ?? 0),
                ),
                "svg:y2": emuToLength(
                  odfLengthToEmu(shape.rowOffset ?? 0) + odfLengthToEmu(shape.extentCy ?? 0),
                ),
              }
            : {
                "svg:x": emuToLength(shape.colOffset ?? 0),
                "svg:y": emuToLength(shape.rowOffset ?? 0),
                "svg:width": emuToLength(shape.extentCx ?? 0),
                "svg:height": emuToLength(shape.extentCy ?? 0),
              }),
        },
        shape.textBody?.text === undefined
          ? []
          : [xmlElement("text:p", undefined, [shape.textBody.text])],
      );
    });
}

export function tableProtectionXml(worksheet: WorksheetOptions): string {
  const protection = worksheet.protection;
  if (!protection?.sheet) return "";
  return xmlElement("loext:table-protection", {
    "loext:select-protected-cells":
      protection.selectLockedCells === undefined ? undefined : !protection.selectLockedCells,
    "loext:select-unprotected-cells":
      protection.selectUnlockedCells === undefined ? undefined : !protection.selectUnlockedCells,
    "loext:insert-rows": protection.insertRows === undefined ? undefined : !protection.insertRows,
    "loext:delete-rows": protection.deleteRows === undefined ? undefined : !protection.deleteRows,
  });
}

export function parseConsolidation(
  body: Element | undefined,
  worksheets: WorksheetOptions[],
): void {
  const consolidation = childNamed(body, "table:consolidation");
  if (!consolidation) return;
  assertElementNames(consolidation, [
    "table:function",
    "table:source-cell-range-addresses",
    "table:target-cell-address",
    "table:use-labels",
    "table:link-to-source-data",
  ]);
  const target = parseSheetAddress(attributeString(consolidation, "table:target-cell-address"));
  const worksheet = worksheets.find((entry) => entry.name === target?.sheet) ?? worksheets[0];
  if (!worksheet) {
    throw unknownOdsElement(
      consolidation,
      CONTENT_PATH,
      "table:target-cell-address",
      "worksheet is missing",
    );
  }
  worksheet.dataConsolidate = {
    function: consolidationFunction(attributeString(consolidation, "table:function")),
    ...(attributeString(consolidation, "table:use-labels") === "row"
      ? { topLabels: true }
      : attributeString(consolidation, "table:use-labels") === "column"
        ? { leftLabels: true }
        : attributeString(consolidation, "table:use-labels") === "both"
          ? { topLabels: true, leftLabels: true }
          : {}),
    link: attributeString(consolidation, "table:link-to-source-data") === "true" || undefined,
    refs: (attributeString(consolidation, "table:source-cell-range-addresses") ?? "")
      .split(/\s+/u)
      .filter(Boolean)
      .map((source) => {
        const address = parseSheetAddress(source);
        return { ref: address?.address, sheet: address?.sheet };
      }),
  };
}

function consolidationFunction(
  value: string | undefined,
): NonNullable<NonNullable<WorksheetOptions["dataConsolidate"]>["function"]> | undefined {
  const functions: Record<
    string,
    NonNullable<NonNullable<WorksheetOptions["dataConsolidate"]>["function"]>
  > = {
    average: "average",
    count: "count",
    countnums: "countNumbers",
    max: "max",
    min: "min",
    product: "product",
    stdev: "standardDeviation",
    stdevp: "standardDeviationPopulation",
    sum: "sum",
    var: "variance",
    varp: "variancePopulation",
  };
  return value === undefined ? undefined : functions[value];
}

export function consolidationXml(worksheets: WorksheetOptions[]): string {
  const entries = worksheets
    .map((worksheet) => ({ worksheet, consolidate: worksheet.dataConsolidate }))
    .filter((entry) => entry.consolidate !== undefined);
  if (!entries.length) return "";
  const { worksheet, consolidate } = entries[0]!;
  const functions = new Map(
    Object.entries({
      average: "average",
      count: "count",
      countNumbers: "countnums",
      max: "max",
      min: "min",
      product: "product",
      standardDeviation: "stdev",
      standardDeviationPopulation: "stdevp",
      sum: "sum",
      variance: "var",
      variancePopulation: "varp",
    }),
  );
  const labels =
    consolidate!.topLabels && consolidate!.leftLabels
      ? "both"
      : consolidate!.topLabels
        ? "row"
        : consolidate!.leftLabels
          ? "column"
          : "none";
  return xmlElement("table:consolidation", {
    "table:function": functions.get(consolidate!.function ?? "sum"),
    "table:source-cell-range-addresses": (consolidate!.refs ?? [])
      .map((reference) => {
        const source = typeof reference === "string" ? { ref: reference } : reference;
        return source.sheet ? sheetAddress(source.sheet, source.ref ?? "") : source.ref;
      })
      .join(" "),
    "table:target-cell-address": sheetAddress(
      worksheet.name ?? "Sheet1",
      worksheets.indexOf(worksheet) === -1 ? "A1" : "A1",
    ),
    "table:use-labels": labels === "none" ? undefined : labels,
    "table:link-to-source-data": consolidate!.link,
  });
}

export function parseSheetAddress(value: string | undefined): SheetAddress | undefined {
  if (value === undefined) return undefined;
  const match = /^(?:'((?:[^']|'')*)'|([^.]+))\.(.+)$/u.exec(value);
  if (!match) return { address: value };
  return {
    sheet: match[1] !== undefined ? match[1].replace(/''/gu, "'") : match[2],
    address: match[3] ?? value,
  };
}

export function sheetAddress(sheet: string, address: string): string {
  const quoted = `'${sheet.replace(/'/gu, "''")}'`;
  return `${quoted}.${address}`;
}

export function parseCommonCellStyles(styles: Element | undefined): Map<string, StyleOptions> {
  const container = childNamed(styles, "office:styles");
  const result = parseNumberStyles(container);
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-cell") continue;
    const name = attributeString(style, "style:name") ?? "";
    if (!result.has(name)) result.set(name, {});
    const displayName = attributeString(style, "style:display-name");
    if (displayName && !result.has(displayName)) result.set(displayName, result.get(name) ?? {});
  }
  return result;
}

export function parseAutomaticCellStyleNames(
  styles: Element | undefined,
): Map<string, StyleOptions> {
  const result = new Map<string, StyleOptions>();
  for (const style of childrenNamed(styles, "style:style")) {
    if (attributeString(style, "style:family") !== "table-cell") continue;
    result.set(attributeString(style, "style:name") ?? "", {});
  }
  return result;
}

export function parseWorksheetNamedExpressions(
  table: Element,
  sheetIndex: number,
): NonNullable<WorkbookOptions["definedNames"]> {
  const expressions = childNamed(table, "table:named-expressions");
  if (!expressions) return [];
  return [
    ...childrenNamed(expressions, "table:named-range").map((range) => ({
      name: attributeString(range, "table:name") ?? "",
      value: attributeString(range, "table:cell-range-address") ?? "",
      localSheetId: sheetIndex,
      hidden: attributeString(range, "loext:hidden") === "true" || undefined,
    })),
    ...childrenNamed(expressions, "table:named-expression").map((expression) => ({
      name: attributeString(expression, "table:name") ?? "",
      value: attributeString(expression, "table:expression") ?? "",
      localSheetId: sheetIndex,
      hidden: attributeString(expression, "loext:hidden") === "true" || undefined,
    })),
  ];
}

export function namedExpressionsXml(
  definedNames: NonNullable<WorkbookOptions["definedNames"]>,
): string {
  if (!definedNames?.length) return "";
  return xmlElement(
    "table:named-expressions",
    undefined,
    definedNames.map((definedName) =>
      /^[A-Za-z0-9_]+![A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$/.test(definedName.value)
        ? xmlElement("table:named-range", {
            "table:name": definedName.name,
            "table:cell-range-address": definedName.value,
            "loext:hidden": definedName.hidden,
          })
        : xmlElement("table:named-expression", {
            "table:name": definedName.name,
            "table:expression": definedName.value,
            "loext:hidden": definedName.hidden,
          }),
    ),
  );
}

export function parseContentValidations(
  body: Element | undefined,
): Map<string, DataValidationOptions> {
  const result = new Map<string, DataValidationOptions>();
  const container = childNamed(body, "table:content-validations");
  if (!container) return result;
  for (const validation of childrenNamed(container, "table:content-validation")) {
    assertElementNames(
      validation,
      [
        "table:name",
        "table:condition",
        "table:base-cell-address",
        "table:allow-empty-cell",
        "table:display-list",
      ],
      ["table:help-message", "table:error-message", "table:error-macro", "office:event-listeners"],
    );
    const help = childNamed(validation, "table:help-message");
    const error = childNamed(validation, "table:error-message");
    if (help) assertElementNames(help, ["table:display", "table:title"], ["text:p"]);
    if (error)
      assertElementNames(error, ["table:display", "table:title", "table:message-type"], ["text:p"]);
    const condition = attributeString(validation, "table:condition");
    const displayList = attributeString(validation, "table:display-list");
    result.set(attributeString(validation, "table:name") ?? "", {
      sqref: "",
      ...(condition ? { type: "custom", formula1: condition } : {}),
      allowBlank: attributeString(validation, "table:allow-empty-cell") !== "false",
      showErrorMessage:
        error === undefined ? undefined : attributeString(error, "table:display") === "true",
      errorTitle: attributeString(error, "table:title"),
      error:
        error === undefined ? undefined : childrenNamed(error, "text:p").map(textOf).join("\n"),
      errorStyle:
        attributeString(error, "table:message-type") === "warning"
          ? "warning"
          : attributeString(error, "table:message-type") === "macro"
            ? "information"
            : attributeString(error, "table:message-type") === undefined
              ? undefined
              : "stop",
      showInputMessage:
        help === undefined ? undefined : attributeString(help, "table:display") === "true",
      promptTitle: attributeString(help, "table:title"),
      prompt: help === undefined ? undefined : childrenNamed(help, "text:p").map(textOf).join("\n"),
      showDropDown: displayList === "none" ? true : displayList === undefined ? undefined : false,
    });
  }
  return result;
}

export function attachContentValidations(
  table: Element,
  worksheet: WorksheetOptions,
  declarations: Map<string, DataValidationOptions>,
): DataValidationOptions[] {
  const attached = new Map<string, string[]>();
  let rowNumber = 1;
  for (const row of tableRows(table)) {
    const rowRepeat = attributeNumber(row, "table:number-rows-repeated") ?? 1;
    let columnNumber = 1;
    for (const cell of row.elements ?? []) {
      if (cell.type !== "element" || !cell.name) continue;
      const repeat = attributeNumber(cell, "table:number-columns-repeated") ?? 1;
      const name = attributeString(cell, "table:content-validation-name");
      if (name) {
        const column = spreadsheetColumnNumber(`${columnNameText(columnNumber)}`);
        attached.set(name, [
          ...(attached.get(name) ?? []),
          `${columnNameText(column)}${rowNumber}`,
        ]);
        void column;
      }
      columnNumber += repeat;
    }
    rowNumber += rowRepeat;
  }
  const result: DataValidationOptions[] = [];
  for (const [name, references] of attached) {
    const declaration = declarations.get(name);
    if (!declaration) {
      throw unknownOdsElement(
        table,
        CONTENT_PATH,
        name,
        "content validation declaration is missing",
      );
    }
    result.push({ ...declaration, sqref: references.join(" ") });
  }
  worksheet.dataValidations = result.length ? result : undefined;
  return result;
}

function columnNameText(column: number): string {
  let value = column;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function tableRows(element: Element): Element[] {
  return (element.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-row") return [child];
    if (
      child.name === "table:table-rows" ||
      child.name === "table:table-header-rows" ||
      child.name === "table:table-row-group"
    ) {
      return tableRows(child);
    }
    return [];
  });
}

export function contentValidationsXml(worksheets: WorksheetOptions[]): string {
  const declarations = worksheets.flatMap((worksheet, worksheetIndex) =>
    (worksheet.dataValidations ?? []).map((validation, validationIndex) => ({
      validation,
      name: `val${worksheetIndex + 1}_${validationIndex + 1}`,
      worksheet,
    })),
  );
  if (!declarations.length) return "";
  return xmlElement(
    "table:content-validations",
    undefined,
    declarations.map(({ validation, name, worksheet }) =>
      xmlElement(
        "table:content-validation",
        {
          "table:name": name,
          "table:condition": validation.formula1,
          "table:base-cell-address": sheetAddress(
            worksheet.name ?? "Sheet1",
            validation.sqref.split(/\s+/u)[0] ?? "A1",
          ),
          "table:allow-empty-cell": validation.allowBlank === false ? "false" : "true",
          "table:display-list":
            validation.showDropDown === true
              ? "none"
              : validation.showDropDown === false
                ? "unsorted"
                : undefined,
        },
        [
          ...(validation.promptTitle !== undefined || validation.prompt !== undefined
            ? [
                xmlElement(
                  "table:help-message",
                  {
                    "table:display": validation.showInputMessage === false ? "false" : "true",
                    "table:title": validation.promptTitle,
                  },
                  paragraphsXml(validation.prompt),
                ),
              ]
            : []),
          ...(validation.errorTitle !== undefined || validation.error !== undefined
            ? [
                xmlElement(
                  "table:error-message",
                  {
                    "table:display": validation.showErrorMessage === false ? "false" : "true",
                    "table:title": validation.errorTitle,
                    "table:message-type":
                      validation.errorStyle === "warning"
                        ? "warning"
                        : validation.errorStyle === "information"
                          ? "macro"
                          : "stop",
                  },
                  paragraphsXml(validation.error),
                ),
              ]
            : []),
        ],
      ),
    ),
  );
}

export function contentValidationNames(worksheets: WorksheetOptions[]): Map<string, string> {
  const result = new Map<string, string>();
  worksheets.forEach((worksheet, worksheetIndex) => {
    (worksheet.dataValidations ?? []).forEach((validation, validationIndex) => {
      const name = `val${worksheetIndex + 1}_${validationIndex + 1}`;
      for (const reference of validation.sqref.split(/\s+/u)) {
        result.set(`${worksheet.name ?? `Sheet${worksheetIndex + 1}`}:${reference}`, name);
      }
    });
  });
  return result;
}

function paragraphsXml(text: string | undefined): string[] {
  if (text === undefined) return [];
  return text.split("\n").map((paragraph) => xmlElement("text:p", undefined, [paragraph]));
}

export function parseDatabaseRanges(
  body: Element | undefined,
  worksheets: WorksheetOptions[],
): void {
  const container = childNamed(body, "table:database-ranges");
  if (!container) return;
  for (const range of childrenNamed(container, "table:database-range")) {
    assertElementNames(
      range,
      [
        "table:name",
        "table:is-selection",
        "table:on-update-keep-styles",
        "table:on-update-keep-size",
        "table:has-persistent-data",
        "table:orientation",
        "table:contains-header",
        "table:display-filter-buttons",
        "table:target-range-address",
        "calcext:contains-footer",
      ],
      [
        "table:database-source-sql",
        "table:database-source-table",
        "table:database-source-query",
        "table:filter",
        "table:sort",
        "table:subtotal-rules",
      ],
    );
    const target = parseSheetAddress(attributeString(range, "table:target-range-address"));
    if (!target?.sheet || !target.address) {
      throw unknownOdsElement(
        range,
        CONTENT_PATH,
        "table:target-range-address",
        "database range target is invalid",
      );
    }
    const worksheet = worksheets.find((entry) => entry.name === target.sheet);
    if (!worksheet) {
      throw unknownOdsElement(
        range,
        CONTENT_PATH,
        target.sheet,
        "database range worksheet is missing",
      );
    }
    worksheet.tables = [
      ...(worksheet.tables ?? []),
      {
        displayName:
          attributeString(range, "table:name") ??
          `DatabaseRange${worksheet.tables?.length ?? 0 + 1}`,
        name: attributeString(range, "table:name"),
        ref: target.address,
        columns: [],
        headerRowCount: attributeString(range, "table:contains-header") === "false" ? 0 : undefined,
      },
    ];
  }
}

export function databaseRangesXml(worksheets: WorksheetOptions[]): string {
  const ranges = worksheets.flatMap((worksheet) =>
    (worksheet.tables ?? []).map((table) => ({ worksheet, table })),
  );
  if (!ranges.length) return "";
  return xmlElement(
    "table:database-ranges",
    undefined,
    ranges.map(({ worksheet, table }) =>
      xmlElement(
        "table:database-range",
        {
          "table:name": table.displayName,
          "table:target-range-address": sheetAddress(worksheet.name ?? "Sheet1", table.ref),
          "table:contains-header": table.headerRowCount === 0 ? "false" : undefined,
          "table:display-filter-buttons": table.autoFilter === undefined ? undefined : true,
        },
        table.autoFilter === undefined || typeof table.autoFilter === "string"
          ? []
          : [filterXml(table)],
      ),
    ),
  );
}

export function dataPilotTablesXml(worksheets: WorksheetOptions[]): string {
  const tables = worksheets.flatMap((worksheet) =>
    (worksheet.pivotTables ?? [])
      .filter((pivot) => pivot.mode === "source")
      .map((pivot) => ({ worksheet, pivot })),
  );
  if (!tables.length) return "";
  return xmlElement(
    "table:data-pilot-tables",
    undefined,
    tables.map(({ worksheet, pivot }, index) =>
      xmlElement(
        "table:data-pilot-table",
        {
          "table:name": pivot.name ?? `DataPilot${index + 1}`,
          "table:target-range-address": sheetAddress(
            worksheet.name ?? `Sheet${index + 1}`,
            `${pivot.location}:${pivot.location}`,
          ),
          "table:grand-total": "none",
          "table:show-filter-button": false,
          "table:drill-down-on-double-click": false,
        },
        [
          xmlElement("table:source-cell-range", {
            "table:cell-range-address": sheetAddress(
              pivot.sourceSheet ?? worksheet.name ?? "Sheet1",
              pivot.source,
            ),
          }),
          ...pivot.rows.map((field) =>
            xmlElement("table:data-pilot-field", {
              "table:source-field-name": field,
              "table:orientation": "row",
              "table:function": "auto",
            }),
          ),
          ...(pivot.columns ?? []).map((field) =>
            xmlElement("table:data-pilot-field", {
              "table:source-field-name": field,
              "table:orientation": "column",
              "table:function": "auto",
            }),
          ),
          ...pivot.data.map((field) =>
            xmlElement("table:data-pilot-field", {
              "table:source-field-name": field.field,
              "table:orientation": "data",
              "table:function": field.summarize ?? "auto",
            }),
          ),
        ],
      ),
    ),
  );
}

function filterXml(table: TableOptions): string {
  const autoFilter = typeof table.autoFilter === "string" ? undefined : table.autoFilter;
  const conditions = autoFilter?.columns ?? [];
  if (!conditions.length) return "";
  const group = conditions.length > 1 ? "table:filter-or" : undefined;
  const children = conditions.map((column) => {
    const values = column.filters?.values ?? [];
    return xmlElement(
      "table:filter-condition",
      {
        "table:field-number": column.colId,
        "table:value": values[0] ?? "",
        "table:operator": "=",
      },
      values.map((value) => xmlElement("table:filter-set-item", { "table:value": value })),
    );
  });
  return xmlElement(
    "table:filter",
    undefined,
    group ? [xmlElement(group, undefined, children)] : children,
  );
}

export function parseConditionalFormats(
  table: Element,
  sheetName: string,
  styles: Map<string, StyleOptions>,
  dxfs: DxfOptions[],
): ConditionalFormatOptions[] {
  const container = childNamed(table, "calcext:conditional-formats");
  if (!container) return [];
  const result: ConditionalFormatOptions[] = [];
  for (const format of childrenNamed(container, "calcext:conditional-format")) {
    assertElementNames(
      format,
      ["calcext:target-range-address"],
      [
        "calcext:condition",
        "calcext:color-scale",
        "calcext:data-bar",
        "calcext:icon-set",
        "calcext:date-is",
      ],
    );
    const target = parseSheetAddress(attributeString(format, "calcext:target-range-address"));
    const rules: ConditionalFormatRule[] = childrenNamed(format, "calcext:condition").map(
      (condition) => {
        assertElementNames(condition, [
          "calcext:apply-style-name",
          "calcext:base-cell-address",
          "calcext:value",
        ]);
        const styleName = attributeString(condition, "calcext:apply-style-name") ?? "";
        const style = styles.get(styleName);
        if (!style) {
          throw unknownOdsElement(
            condition,
            CONTENT_PATH,
            styleName,
            "conditional style is missing",
          );
        }
        return {
          ...conditionalRule(attributeString(condition, "calcext:value") ?? ""),
          priority: result.length + rulesPrioritySeed(format),
          dxfId: dxfIndex(style, dxfs),
        };
      },
    );
    const colorScale = childNamed(format, "calcext:color-scale");
    if (colorScale) rules.push(colorScaleRule(colorScale));
    const dataBar = childNamed(format, "calcext:data-bar");
    if (dataBar) rules.push(dataBarRule(dataBar));
    const iconSet = childNamed(format, "calcext:icon-set");
    if (iconSet) rules.push(iconSetRule(iconSet));
    if (!rules.length) continue;
    result.push({ sqref: target?.address ?? sheetName, rules });
  }
  return result;
}

export function conditionalFormatsXml(
  formats: ConditionalFormatOptions[],
  sheetName: string,
  styleNames: Map<number, string>,
): string {
  if (!formats.length) return "";
  return xmlElement(
    "calcext:conditional-formats",
    undefined,
    formats.flatMap((format) => {
      const rules = format.rules.map((rule) => {
        if (rule.type === "colorScale") {
          return xmlElement(
            "calcext:color-scale",
            undefined,
            rule.colorScale?.cfvo.map((entry, entryIndex) =>
              xmlElement("calcext:color-scale-entry", {
                "calcext:type": cfvoType(entry.type),
                "calcext:value": entry.val,
                "calcext:color": `#${rule.colorScale?.colors[entryIndex]?.rgb?.slice(2) ?? "000000"}`,
              }),
            ) ?? [],
          );
        }
        if (rule.type === "dataBar") {
          return xmlElement(
            "calcext:data-bar",
            {
              "calcext:positive-color": `#${rule.dataBar?.color.rgb?.slice(2) ?? "0000FF"}`,
              "calcext:max-length": rule.dataBar?.maxLength,
            },
            rule.dataBar?.cfvo.map((entry) =>
              xmlElement("calcext:formatting-entry", {
                "calcext:type": cfvoType(entry.type),
                "calcext:value": entry.val,
              }),
            ) ?? [],
          );
        }
        if (rule.type === "iconSet") {
          return xmlElement(
            "calcext:icon-set",
            { "calcext:icon-set-type": rule.iconSet?.iconSet ?? "3TrafficLights1" },
            rule.iconSet?.cfvo.map((entry) =>
              xmlElement("calcext:formatting-entry", {
                "calcext:type": cfvoType(entry.type),
                "calcext:value": entry.val,
              }),
            ) ?? [],
          );
        }
        return xmlElement("calcext:condition", {
          "calcext:apply-style-name": styleNames.get(rule.dxfId ?? -1) ?? "Default",
          "calcext:value": conditionValue(rule),
          "calcext:base-cell-address": sheetAddress(
            sheetName,
            format.sqref.split(/[\s:]/u)[0] ?? "A1",
          ),
        });
      });
      if (!rules.length) return [];
      return [
        xmlElement(
          "calcext:conditional-format",
          { "calcext:target-range-address": sheetAddress(sheetName, format.sqref) },
          rules,
        ),
      ];
    }),
  );
}

function cfvoType(type: CfvoOptions["type"]): string {
  return type === "min" ? "minimum" : type === "max" ? "maximum" : type === "num" ? "number" : type;
}

function conditionValue(rule: ConditionalFormatRule): string {
  if (rule.type === "expression") return `formula-is(${rule.formulas?.[0] ?? ""})`;
  if (rule.type === "top10") {
    const prefix = rule.bottom ? "bottom" : "top";
    return `${prefix}-${rule.percent ? "percent" : "elements"}(${rule.rank ?? 0})`;
  }
  if (rule.type === "aboveAverage")
    return rule.aboveAverage === false ? "below-average" : "above-average";
  if (rule.type === "containsText" && rule.operator === "beginsWith") {
    return `begins-with(${rule.text ?? ""})`;
  }
  const operator =
    rule.operator === "notEqual"
      ? "!="
      : rule.operator === "lessThanOrEqual"
        ? "<="
        : rule.operator === "greaterThanOrEqual"
          ? ">="
          : rule.operator === "equal"
            ? "="
            : rule.operator === "lessThan"
              ? "<"
              : rule.operator === "greaterThan"
                ? ">"
                : "=";
  if (rule.operator === "between" || rule.operator === "notBetween") {
    return `${rule.operator === "between" ? "" : "not-"}between(${rule.formulas?.join(",") ?? ""})`;
  }
  return `${operator}${rule.formulas?.[0] ?? ""}`;
}

function rulesPrioritySeed(format: Element): number {
  return (format.elements ?? []).filter((child) => child.type === "element").length;
}

function dxfIndex(style: StyleOptions, dxfs: DxfOptions[]): number {
  const dxf: DxfOptions = {
    font: style.font,
    fill: style.fill,
    border: style.border,
    numFmt: style.numFmt,
    alignment: style.alignment,
    protection: style.protection,
  };
  const index = dxfs.findIndex((entry) => JSON.stringify(entry) === JSON.stringify(dxf));
  if (index >= 0) return index;
  dxfs.push(dxf);
  return dxfs.length - 1;
}

function conditionalRule(value: string): ConditionalFormatRule {
  if (value.startsWith("formula-is(")) {
    return { type: "expression", formulas: [value.slice("formula-is(".length, -1)] };
  }
  if (value.startsWith("between(") || value.startsWith("not-between(")) {
    const formulas = splitArguments(value.slice(value.indexOf("(") + 1, -1));
    return {
      type: "cellIs",
      operator: value.startsWith("between(") ? "between" : "notBetween",
      formulas,
    };
  }
  if (value.startsWith("top-elements(") || value.startsWith("bottom-elements(")) {
    return {
      type: "top10",
      rank: Number(value.slice(value.indexOf("(") + 1, -1)),
      bottom: value.startsWith("bottom-elements(") || undefined,
    };
  }
  if (value.startsWith("top-percent(") || value.startsWith("bottom-percent(")) {
    return {
      type: "top10",
      percent: true,
      rank: Number(value.slice(value.indexOf("(") + 1, -1)),
      bottom: value.startsWith("bottom-percent(") || undefined,
    };
  }
  if (value === "above-average" || value === "below-average") {
    return { type: "aboveAverage", aboveAverage: value === "above-average" };
  }
  if (value.startsWith("contains-text(")) {
    return {
      type: "containsText",
      operator: "containsText",
      text: value.slice("contains-text(".length, -1),
    };
  }
  if (value.startsWith("begins-with(")) {
    return {
      type: "containsText",
      operator: "beginsWith",
      text: value.slice("begins-with(".length, -1),
    };
  }
  if (value === "duplicate") return { type: "expression", formulas: ["COUNTIF(ROW(),ROW())>1"] };
  if (value === "unique") return { type: "expression", formulas: ["COUNTIF(ROW(),ROW())=1"] };
  const comparison = /^(<=|>=|!=|=|<|>)(.+)$/u.exec(value);
  if (comparison) {
    return {
      type: "cellIs",
      operator:
        comparison[1] === "!="
          ? "notEqual"
          : comparison[1] === "<="
            ? "lessThanOrEqual"
            : comparison[1] === ">="
              ? "greaterThanOrEqual"
              : comparison[1] === "="
                ? "equal"
                : comparison[1] === "<"
                  ? "lessThan"
                  : "greaterThan",
      formulas: [comparison[2] ?? ""],
    };
  }
  return { type: "expression", formulas: [value] };
}

function splitArguments(value: string): string[] {
  return value.split(/,(?=(?:[^()*]*\([^()]*\))*[^()]*$)/u);
}

function cfvo(element: Element): CfvoOptions {
  const type = attributeString(element, "calcext:type") ?? "number";
  return {
    type:
      type === "minimum" || type === "maximum" || type === "auto-minimum" || type === "auto-maximum"
        ? type.startsWith("auto")
          ? type === "auto-minimum"
            ? "min"
            : "max"
          : type === "minimum"
            ? "min"
            : "max"
        : type === "number"
          ? "num"
          : type === "percentile"
            ? "percentile"
            : type === "formula"
              ? "formula"
              : "percent",
    val: attributeNumber(element, "calcext:value") ?? attributeString(element, "calcext:value"),
  };
}

function colorScaleRule(element: Element): ConditionalFormatRule {
  const entries = childrenNamed(element, "calcext:color-scale-entry");
  return {
    type: "colorScale",
    colorScale: {
      cfvo: entries.map(cfvo),
      colors: entries.map((entry) => ({
        rgb: `FF${attributeString(entry, "calcext:color")?.replace("#", "") ?? "000000"}`,
      })),
    },
  };
}

function dataBarRule(element: Element): ConditionalFormatRule {
  const entries = childrenNamed(element, "calcext:formatting-entry");
  return {
    type: "dataBar",
    dataBar: {
      cfvo: [cfvo(entries[0] ?? element), cfvo(entries[1] ?? element)],
      color: {
        rgb: `FF${attributeString(element, "calcext:positive-color")?.replace("#", "") ?? "0000FF"}`,
      },
      maxLength: attributeNumber(element, "calcext:max-length"),
    },
  };
}

const ICON_SET_TYPES = new Set<string>([
  "3Arrows",
  "3ArrowsGray",
  "3Flags",
  "3TrafficLights1",
  "3TrafficLights2",
  "3Signs",
  "3Symbols",
  "3Symbols2",
  "4Arrows",
  "4ArrowsGray",
  "4RedToBlack",
  "4Rating",
  "4TrafficLights",
  "5Arrows",
  "5ArrowsGray",
  "5Rating",
  "5Quarters",
] satisfies IconSetType[]);

function iconSetType(element: Element, value: string | undefined): IconSetType {
  const result = value ?? "3TrafficLights1";
  if (!ICON_SET_TYPES.has(result)) {
    throw unknownOdsElement(element, CONTENT_PATH, result, "icon set type is unsupported");
  }
  return result as IconSetType;
}

function iconSetRule(element: Element): ConditionalFormatRule {
  return {
    type: "iconSet",
    iconSet: {
      cfvo: childrenNamed(element, "calcext:formatting-entry").map(cfvo),
      iconSet: iconSetType(element, attributeString(element, "calcext:icon-set-type")),
    },
  };
}

export function parseDataPilotTables(
  body: Element | undefined,
  worksheets: WorksheetOptions[],
): void {
  const container = childNamed(body, "table:data-pilot-tables");
  if (!container) return;
  for (const table of childrenNamed(container, "table:data-pilot-table")) {
    assertElementNames(
      table,
      [
        "table:name",
        "table:application-data",
        "table:target-range-address",
        "table:buttons",
        "table:grand-total",
        "table:show-filter-button",
        "table:drill-down-on-double-click",
        "table:ignore-empty-rows",
        "table:header-grid-layout",
      ],
      [
        "table:source-cell-range",
        "table:database-source-query",
        "table:data-pilot-field",
        "tableooo:data-pilot-grand-total",
      ],
    );
    const target = parseSheetAddress(attributeString(table, "table:target-range-address"));
    if (!target?.sheet) {
      throw unknownOdsElement(
        table,
        CONTENT_PATH,
        "table:target-range-address",
        "data pilot target is invalid",
      );
    }
    const worksheet = worksheets.find((entry) => entry.name === target.sheet);
    if (!worksheet) {
      throw unknownOdsElement(table, CONTENT_PATH, target.sheet, "data pilot worksheet is missing");
    }
    const sourceRange = childNamed(table, "table:source-cell-range");
    const source = parseSheetAddress(attributeString(sourceRange, "table:cell-range-address"));
    const rows: string[] = [];
    const columns: string[] = [];
    const data: PivotDataField[] = [];
    for (const field of childrenNamed(table, "table:data-pilot-field")) {
      const name = attributeString(field, "table:source-field-name") ?? "";
      const orientation = attributeString(field, "table:orientation");
      if (orientation === "row") rows.push(name);
      else if (orientation === "column") columns.push(name);
      else if (orientation === "data") {
        data.push({
          field: name,
          summarize:
            attributeString(field, "table:function") === "auto"
              ? undefined
              : (attributeString(field, "table:function") as PivotDataField["summarize"]),
        });
      }
    }
    worksheet.pivotTables = [
      ...(worksheet.pivotTables ?? []),
      {
        mode: "source",
        name: attributeString(table, "table:name"),
        source: source?.address ?? "",
        sourceSheet: source?.sheet,
        location: target.address.split(":")[0],
        rows,
        columns,
        data,
      },
    ];
  }
}
