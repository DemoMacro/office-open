import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
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
  AutoFilterOptions,
  CustomFilterOperator,
  DataValidationOptions,
  DxfOptions,
  FilterColumnOptions,
  IconSetType,
  PivotDataField,
  PivotPageFieldOptions,
  SourcePivotTableOptions,
  ShapeOptions,
  SheetProtectionOptions,
  SortCondition,
  SortStateOptions,
  StyleOptions,
  TableOptions,
  WorksheetOptions,
  WorkbookOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { OdsGenerateError } from "./error";
import { parseNumberStyles } from "./numbering";
import type { OdsDataPilotFunctionOptions } from "./options";
import { assertElementNames, unknownOdsElement } from "./workbook";

const CONTENT_PATH = "/office:document-content/office:body/office:spreadsheet";

const ODF_PIVOT_FUNCTIONS = {
  sum: "sum",
  count: "count",
  average: "average",
  max: "max",
  min: "min",
  product: "product",
  countnums: "countNumbers",
  stddev: "standardDeviation",
  stddevp: "standardDeviationPopulation",
  var: "variance",
  varp: "variancePopulation",
} as const;

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
    ...(attributeString(element, "svg:title")
      ? { title: attributeString(element, "svg:title") }
      : {}),
    ...(attributeString(element, "svg:desc")
      ? { description: attributeString(element, "svg:desc") }
      : {}),
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
          ...(shape.title === undefined ? {} : { "svg:title": shape.title }),
          ...(shape.description === undefined ? {} : { "svg:desc": shape.description }),
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
          : [xmlElement("text:p", undefined, [escapeText(shape.textBody.text)])],
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
    "table:use-label",
    "table:link-to-source-data",
  ]);
  const labels =
    attributeString(consolidation, "table:use-labels") ??
    attributeString(consolidation, "table:use-label");
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
    ...(labels === "row"
      ? { topLabels: true }
      : labels === "column"
        ? { leftLabels: true }
        : labels === "both"
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
    for (const name of ["table:error-macro", "office:event-listeners"]) {
      if (childNamed(validation, name)) {
        throw unknownOdsElement(
          validation,
          CONTENT_PATH,
          name,
          "content validation behavior has no canonical DataValidationOptions mapping",
        );
      }
    }
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
  return text
    .split("\n")
    .map((paragraph) => xmlElement("text:p", undefined, [escapeText(paragraph)]));
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
    const sourceNames = [
      "table:database-source-sql",
      "table:database-source-table",
      "table:database-source-query",
    ];
    const source = sourceNames.find((name) => childNamed(range, name) !== undefined);
    if (source) {
      throw unknownOdsElement(
        range,
        CONTENT_PATH,
        source,
        "database source has no canonical TableOptions mapping",
      );
    }
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
    const filter = childNamed(range, "table:filter");
    const sort = childNamed(range, "table:sort");
    const table: TableOptions = {
      displayName:
        attributeString(range, "table:name") ??
        `DatabaseRange${(worksheet.tables?.length ?? 0) + 1}`,
      name: attributeString(range, "table:name"),
      ref: target.address,
      columns: [],
      headerRowCount: parseOdsBoolean(range, "table:contains-header") === false ? 0 : undefined,
      ...(filter ? { autoFilter: parseDatabaseFilter(filter, target.address) } : {}),
      ...(sort ? { sortState: parseDatabaseSort(sort, range, target.address) } : {}),
    };
    const displayFilterButtons = parseOdsBoolean(range, "table:display-filter-buttons");
    if (displayFilterButtons === true && filter === undefined) {
      table.autoFilter = { ref: target.address };
    }
    const unsupportedAttributes = [
      "table:is-selection",
      "table:on-update-keep-styles",
      "table:on-update-keep-size",
      "table:has-persistent-data",
      "table:orientation",
      "calcext:contains-footer",
    ].filter((name) => attributeString(range, name) !== undefined);
    if (unsupportedAttributes.length > 0) {
      throw unknownOdsElement(
        range,
        CONTENT_PATH,
        unsupportedAttributes[0]!,
        "database range attribute has no canonical TableOptions mapping",
      );
    }
    if (childNamed(range, "table:subtotal-rules")) {
      throw unknownOdsElement(
        range,
        CONTENT_PATH,
        "table:subtotal-rules",
        "subtotal rules have no canonical TableOptions mapping",
      );
    }
    worksheet.tables = [...(worksheet.tables ?? []), table];
  }
}

function parseDatabaseFilter(filter: Element, rangeReference: string): AutoFilterOptions {
  assertElementNames(
    filter,
    [
      "table:target-range-address",
      "table:condition-source",
      "table:condition-source-range-address",
      "table:display-duplicates",
    ],
    ["table:filter-condition", "table:filter-and", "table:filter-or"],
  );
  if (
    attributeString(filter, "table:condition-source") !== undefined &&
    attributeString(filter, "table:condition-source") !== "self"
  ) {
    throw unknownOdsElement(
      filter,
      CONTENT_PATH,
      "table:condition-source",
      "external filter condition sources have no canonical autoFilter mapping",
    );
  }
  for (const name of ["table:condition-source-range-address", "table:display-duplicates"]) {
    if (attributeString(filter, name) !== undefined) {
      throw unknownOdsElement(
        filter,
        CONTENT_PATH,
        name,
        "filter attribute has no canonical autoFilter mapping",
      );
    }
  }
  const children = (filter.elements ?? []).filter((child) => child.type === "element");
  const topLevelOr = children.length === 1 && children[0]?.name === "table:filter-or";
  const conditionNodes: { element: Element; group?: "and" | "or" }[] = [];
  const appendGroup = (element: Element, group?: "and" | "or"): void => {
    for (const child of (element.elements ?? []).filter(
      (entry): entry is Element => entry.type === "element",
    )) {
      if (child.name === "table:filter-condition") conditionNodes.push({ element: child, group });
      else if (child.name === "table:filter-or" || child.name === "table:filter-and") {
        appendGroup(child, child.name === "table:filter-or" ? "or" : "and");
      } else {
        throw unknownOdsElement(
          child,
          CONTENT_PATH,
          child.name ?? "table:filter-condition",
          "nested filter logic has no canonical autoFilter mapping",
        );
      }
    }
  };
  for (const child of children) {
    if (child.name === "table:filter-condition") conditionNodes.push({ element: child });
    else if (child.name === "table:filter-or" || child.name === "table:filter-and") {
      appendGroup(child, child.name === "table:filter-or" ? "or" : undefined);
    } else {
      throw unknownOdsElement(
        child,
        CONTENT_PATH,
        child.name ?? "table:filter-condition",
        "nested filter logic has no canonical autoFilter mapping",
      );
    }
  }
  const conditions = conditionNodes.map(({ element: condition, group }) => {
    if (condition.name !== "table:filter-condition") {
      throw unknownOdsElement(
        condition,
        CONTENT_PATH,
        condition.name ?? "table:filter-condition",
        "nested filter logic has no canonical autoFilter mapping",
      );
    }
    assertElementNames(
      condition,
      [
        "table:field-number",
        "table:value",
        "table:operator",
        "table:case-sensitive",
        "table:data-type",
      ],
      ["table:filter-set-item"],
    );
    const field = attributeNumber(condition, "table:field-number");
    const operator = attributeString(condition, "table:operator");
    const values = childrenNamed(condition, "table:filter-set-item").map((item) => {
      assertElementNames(item, ["table:value"]);
      return attributeString(item, "table:value") ?? "";
    });
    const value = attributeString(condition, "table:value");
    if (field === undefined || operator === undefined || value === undefined) {
      throw unknownOdsElement(
        condition,
        CONTENT_PATH,
        "table:filter-condition",
        "required filter data is missing",
      );
    }
    if (values.length > 0 && operator !== "=" && operator !== "equal") {
      throw unknownOdsElement(
        condition,
        CONTENT_PATH,
        "table:filter-set-item",
        "filter set only supports equality",
      );
    }
    return { field, operator, values: values.length > 0 ? values : [value], group };
  });
  const byField = new Map<number, typeof conditions>();
  for (const condition of conditions) {
    const values = byField.get(condition.field) ?? [];
    values.push(condition);
    byField.set(condition.field, values);
  }
  if (topLevelOr && byField.size !== 1) {
    throw unknownOdsElement(
      filter,
      CONTENT_PATH,
      "table:filter-or",
      "OR filters across fields have no canonical autoFilter mapping",
    );
  }
  if (
    [...byField.values()].some(
      (entries) => entries.length > 1 && entries.some((entry) => entry.group === undefined),
    )
  ) {
    throw unknownOdsElement(
      filter,
      CONTENT_PATH,
      "table:filter-condition",
      "multiple AND comparisons for one field have no canonical mapping",
    );
  }
  const columns: FilterColumnOptions[] = [...byField].map(([colId, entries]) => {
    const equality = entries.every((entry) => entry.operator === "=" || entry.operator === "equal");
    if (equality) {
      if (entries.some((entry) => entry.group === "and")) {
        return {
          colId,
          customFilters: {
            and: true,
            entries: entries.map((entry) => ({ operator: "equal", val: entry.values[0]! })),
          },
        };
      }
      return { colId, filters: { values: entries.flatMap((entry) => entry.values) } };
    }
    const mapped: Record<string, CustomFilterOperator> = {
      "!=": "notEqual",
      "<>": "notEqual",
      "not-equal": "notEqual",
      ">": "greaterThan",
      ">=": "greaterThanOrEqual",
      "<": "lessThan",
      "<=": "lessThanOrEqual",
    };
    const customEntries: {
      operator: CustomFilterOperator;
      value: string;
      group: "and" | "or" | undefined;
    }[] = [];
    for (const entry of entries) {
      const canonicalOperator = mapped[entry.operator];
      const value = entry.values[0];
      if (canonicalOperator === undefined || value === undefined || entry.group === undefined) {
        throw unknownOdsElement(
          filter,
          CONTENT_PATH,
          "table:filter-condition",
          `unsupported filter operator: ${entry.operator}`,
        );
      }
      customEntries.push({ operator: canonicalOperator, value, group: entry.group });
    }
    return {
      colId,
      customFilters: {
        ...(customEntries.every((entry) => entry.group === "and") ? { and: true } : {}),
        entries: customEntries.map((entry) => ({ operator: entry.operator, val: entry.value })),
      },
    };
  });
  return {
    ref:
      parseSheetAddress(attributeString(filter, "table:target-range-address"))?.address ??
      rangeReference,
    columns,
  };
}

function parseDatabaseSort(
  sort: Element,
  range: Element,
  rangeReference: string,
): SortStateOptions {
  assertElementNames(
    sort,
    [
      "table:bind-styles-to-content",
      "table:target-range-address",
      "table:case-sensitive",
      "table:language",
      "table:country",
      "table:script",
      "table:rfc-language-tag",
      "table:algorithm",
      "table:embedded-number-behavior",
    ],
    ["table:sort-by"],
  );
  for (const name of [
    "table:bind-styles-to-content",
    "table:language",
    "table:country",
    "table:script",
    "table:rfc-language-tag",
    "table:algorithm",
    "table:embedded-number-behavior",
  ]) {
    if (attributeString(sort, name) !== undefined) {
      throw unknownOdsElement(
        sort,
        CONTENT_PATH,
        name,
        "sort attribute has no canonical sortState mapping",
      );
    }
  }
  const reference =
    parseSheetAddress(attributeString(sort, "table:target-range-address"))?.address ??
    rangeReference;
  const rangeStart = /^\$?([A-Z]+)\$?(\d+)/u.exec(reference.split(":")[0] ?? "");
  if (!rangeStart) {
    throw unknownOdsElement(
      range,
      CONTENT_PATH,
      "table:target-range-address",
      "database range has an invalid sort reference",
    );
  }
  const firstColumn = spreadsheetColumnNumber(rangeStart[1]!);
  const firstRow = Number(rangeStart[2]);
  const conditions: SortCondition[] = childrenNamed(sort, "table:sort-by").map((entry) => {
    assertElementNames(entry, ["table:field-number", "table:data-type", "table:order"]);
    const field = attributeNumber(entry, "table:field-number");
    if (field === undefined) {
      throw unknownOdsElement(
        entry,
        CONTENT_PATH,
        "table:field-number",
        "sort field number is missing",
      );
    }
    if (attributeString(entry, "table:data-type") !== undefined) {
      throw unknownOdsElement(
        entry,
        CONTENT_PATH,
        "table:data-type",
        "sort data type has no canonical sortState mapping",
      );
    }
    const order = attributeString(entry, "table:order");
    if (order !== undefined && order !== "ascending" && order !== "descending") {
      throw unknownOdsElement(entry, CONTENT_PATH, "table:order", `invalid sort order: ${order}`);
    }
    return {
      ref: `${spreadsheetColumnName(firstColumn + field)}${firstRow}`,
      ...(order === undefined ? {} : { descending: order === "descending" }),
    };
  });
  return {
    ref: reference,
    ...(attributeString(sort, "table:case-sensitive") === undefined
      ? {}
      : { caseSensitive: parseOdsBoolean(sort, "table:case-sensitive") }),
    conditions,
  };
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
        [
          ...(table.autoFilter === undefined || typeof table.autoFilter === "string"
            ? []
            : [filterXml(table, worksheet.name ?? "Sheet1")]),
          ...(table.sortState === undefined ? [] : [sortXml(table, worksheet.name ?? "Sheet1")]),
        ],
      ),
    ),
  );
}

export function dataPilotTablesXml(
  worksheets: WorksheetOptions[],
  functions: OdsDataPilotFunctionOptions[] = [],
): string {
  const overrides = new Map(
    functions.map((item) => [`${item.tableName ?? ""}\0${item.field}`, item.summaryFunction]),
  );
  const tables = worksheets.flatMap((worksheet) =>
    (worksheet.pivotTables ?? [])
      .filter((pivot) => pivot.mode === "source")
      .map((pivot) => ({ worksheet, pivot })),
  );
  if (!tables.length) return "";
  return xmlElement(
    "table:data-pilot-tables",
    undefined,
    tables.map(({ worksheet, pivot }, index) => {
      const sourceSheetName = pivot.sourceSheet ?? worksheet.name ?? `Sheet${index + 1}`;
      const sourceWorksheet = pivot.sourceSheet
        ? worksheets.find((item) => item.name === pivot.sourceSheet)
        : worksheet;
      if (!sourceWorksheet) {
        throw new OdsGenerateError(
          `ODS Data Pilot source sheet does not exist: ${sourceSheetName ?? ""}`,
          "content.xml",
          `/office:document-content/office:body/office:spreadsheet/table:data-pilot-tables/table:data-pilot-table[@table:name="${pivot.name ?? ""}"]`,
          "table:source-cell-range",
          "source-sheet-missing",
        );
      }
      const location = pivot.location ?? "A3";
      return xmlElement(
        "table:data-pilot-table",
        {
          "table:name": pivot.name ?? `DataPilot${index + 1}`,
          "table:target-range-address": sheetAddress(
            worksheet.name ?? `Sheet${index + 1}`,
            `${location}:${location}`,
          ),
          "table:grand-total":
            pivot.rowGrandTotals === false && pivot.colGrandTotals === false
              ? "none"
              : pivot.rowGrandTotals === false
                ? "column"
                : pivot.colGrandTotals === false
                  ? "row"
                  : "both",
          "table:drill-down-on-double-click": pivot.enableDrill,
        },
        [
          xmlElement("table:source-cell-range", {
            "table:cell-range-address": sheetAddress(sourceSheetName, pivot.source),
          }),
          ...pivot.rows.map((field) =>
            xmlElement("table:data-pilot-field", {
              "table:source-field-name": field,
              "table:orientation": "row",
              "table:function": "auto",
            }),
          ),
          ...(pivot.pages ?? []).map((field) =>
            xmlElement("table:data-pilot-field", {
              "table:source-field-name": field.field,
              "table:orientation": "page",
              "table:selected-page":
                field.item === undefined
                  ? undefined
                  : (pivotFieldMembers(sourceWorksheet, pivot.source, field.field)[field.item] ??
                    invalidPageItem(pivot, field.field, field.item)),
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
              "table:function":
                overrides.get(`${pivot.name ?? ""}\0${field.field}`) ?? field.summarize ?? "auto",
            }),
          ),
          ...(pivot.hiddenFields ?? []).map((field) =>
            xmlElement("table:data-pilot-field", {
              "table:source-field-name": field,
              "table:orientation": "hidden",
              "table:function": "auto",
            }),
          ),
        ],
      );
    }),
  );
}

function pivotFieldMembers(worksheet: WorksheetOptions, source: string, field: string): string[] {
  const match = /^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/u.exec(source);
  if (!match) return [];
  const firstColumn = spreadsheetColumnNumber(match[1] ?? "A");
  const firstRow = Number(match[2] ?? 1);
  const lastColumn = spreadsheetColumnNumber(match[3] ?? match[1] ?? "A");
  const lastRow = Number(match[4] ?? match[2] ?? 1);
  const cellAt = (rowNumber: number, column: number) =>
    worksheet.rows
      ?.find((row, index) => (row.rowNumber ?? index + 1) === rowNumber)
      ?.cells?.find((cell, index) => {
        const reference = /^([A-Z]+)[0-9]+$/u.exec(cell.reference ?? "");
        return reference
          ? spreadsheetColumnNumber(reference[1] ?? "A") === column
          : index + 1 === column;
      });
  const fieldColumn = Array.from(
    { length: lastColumn - firstColumn + 1 },
    (_, index) => firstColumn + index,
  ).find((column) => pivotValueText(cellAt(firstRow, column)?.value) === field);
  if (fieldColumn === undefined) return [];
  return [
    ...new Set(
      Array.from({ length: lastRow - firstRow }, (_, index) =>
        pivotValueText(cellAt(firstRow + index + 1, fieldColumn)?.value),
      ),
    ),
  ];
}

function pivotValueText(value: unknown): string {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean")
    return String(value);
  if (value instanceof Date) return value.toISOString();
  return "";
}

function invalidPageItem(pivot: SourcePivotTableOptions, field: string, item: number): never {
  throw new OdsGenerateError(
    `ODS Data Pilot page field ${field} refers to missing source item ${item}`,
    "content.xml",
    `/office:document-content/office:body/office:spreadsheet/table:data-pilot-tables/table:data-pilot-table[@table:name="${pivot.name ?? ""}"]`,
    "table:data-pilot-field",
    "page-field-item-missing",
  );
}

function filterXml(table: TableOptions, worksheetName: string): string {
  const autoFilter = typeof table.autoFilter === "string" ? undefined : table.autoFilter;
  const conditions = autoFilter?.columns ?? [];
  if (!conditions.length) return "";
  const groups = conditions.map((column) => {
    if (column.hiddenButton || column.showButton === false) {
      throw new OdsGenerateError(
        "ODS database filters cannot represent per-column button visibility",
        "content.xml",
        "/office:document-content/office:body/office:spreadsheet/table:database-ranges/table:database-range/table:filter",
        "table:filter-column",
        "filter-button-visibility-unsupported",
      );
    }
    const values = column.filters?.values ?? [];
    const customFilters = column.customFilters?.entries ?? [];
    if (column.filters && column.customFilters) {
      throw new OdsGenerateError(
        "ODS database filters cannot combine value and custom filters for one column",
        "content.xml",
        "/office:document-content/office:body/office:spreadsheet/table:database-ranges/table:database-range/table:filter",
        "table:filter-column",
        "filter-variants-conflict",
      );
    }
    if (customFilters.length > 0) {
      const entries = customFilters.map((entry) =>
        xmlElement("table:filter-condition", {
          "table:field-number": column.colId,
          "table:value": entry.val,
          "table:operator": odsFilterOperator(entry.operator),
        }),
      );
      const group = column.customFilters?.and ? "table:filter-and" : "table:filter-or";
      return entries.length === 1 ? entries : [xmlElement(group, undefined, entries)];
    }
    if (values.length === 0 && !column.filters?.blank) {
      throw new OdsGenerateError(
        "ODS database filter column has no filter values",
        "content.xml",
        "/office:document-content/office:body/office:spreadsheet/table:database-ranges/table:database-range/table:filter",
        "table:filter-column",
        "filter-values-missing",
      );
    }
    const entries = values.length
      ? values.map((value) =>
          xmlElement("table:filter-condition", {
            "table:field-number": column.colId,
            "table:value": value,
            "table:operator": "=",
          }),
        )
      : [
          xmlElement("table:filter-condition", {
            "table:field-number": column.colId,
            "table:value": "",
            "table:operator": "=",
          }),
        ];
    return entries.length === 1 ? entries : [xmlElement("table:filter-or", undefined, entries)];
  });
  const children = groups.flat();
  return xmlElement(
    "table:filter",
    { "table:target-range-address": sheetAddress(worksheetName, autoFilter?.ref ?? table.ref) },
    children.length === 1 ? children : [xmlElement("table:filter-and", undefined, children)],
  );
}

function odsFilterOperator(operator: CustomFilterOperator | undefined): string {
  switch (operator) {
    case "notEqual":
      return "<>";
    case "greaterThan":
      return ">";
    case "greaterThanOrEqual":
      return ">=";
    case "lessThan":
      return "<";
    case "lessThanOrEqual":
      return "<=";
    case "equal":
    case undefined:
      return "=";
  }
}

function sortXml(table: TableOptions, worksheetName: string): string {
  const sortState = table.sortState!;
  const rangeStart = /^([A-Z]+)\d+/u.exec(sortState.ref || table.ref);
  if (!rangeStart) {
    throw new OdsGenerateError(
      "ODS database sort range is invalid",
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/table:database-ranges/table:database-range/table:sort",
      "table:target-range-address",
      "sort-range-invalid",
    );
  }
  const firstColumn = spreadsheetColumnNumber(rangeStart[1]!);
  const sortStart = /^([A-Z]+)\d+/u;
  return xmlElement(
    "table:sort",
    {
      "table:target-range-address": sheetAddress(worksheetName, sortState.ref),
      "table:case-sensitive": sortState.caseSensitive,
    },
    sortState.conditions.map((condition) => {
      const match = sortStart.exec(condition.ref);
      if (!match) {
        throw new OdsGenerateError(
          "ODS database sort condition reference is invalid",
          "content.xml",
          "/office:document-content/office:body/office:spreadsheet/table:database-ranges/table:database-range/table:sort/table:sort-by",
          "table:sort-by",
          "sort-condition-invalid",
        );
      }
      return xmlElement("table:sort-by", {
        "table:field-number": spreadsheetColumnNumber(match[1]!) - firstColumn,
        "table:order": condition.descending ? "descending" : "ascending",
      });
    }),
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
        const style: StyleOptions = styles.get(styleName) ?? {};
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

const dxfKeyCache = new WeakMap<DxfOptions[], string[]>();

function dxfIndex(style: StyleOptions, dxfs: DxfOptions[]): number {
  const dxf: DxfOptions = {
    font: style.font,
    fill: style.fill,
    border: style.border,
    numFmt: style.numFmt,
    alignment: style.alignment,
    protection: style.protection,
  };
  let dxfKeys = dxfKeyCache.get(dxfs);
  if (!dxfKeys) {
    dxfKeys = dxfs.map((entry) => JSON.stringify(entry));
    dxfKeyCache.set(dxfs, dxfKeys);
  }
  const dxfKey = JSON.stringify(dxf);
  const index = dxfKeys.indexOf(dxfKey);
  if (index >= 0) return index;
  dxfs.push(dxf);
  dxfKeys.push(dxfKey);
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
): OdsDataPilotFunctionOptions[] {
  const container = childNamed(body, "table:data-pilot-tables");
  if (!container) return [];
  const functions: OdsDataPilotFunctionOptions[] = [];
  for (const table of childrenNamed(container, "table:data-pilot-table")) {
    assertElementNames(
      table,
      [
        "table:name",
        "table:target-range-address",
        "table:grand-total",
        "table:drill-down-on-double-click",
      ],
      ["table:source-cell-range", "table:data-pilot-field"],
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
    if (!sourceRange) {
      throw unknownOdsElement(
        table,
        CONTENT_PATH,
        "table:source-cell-range",
        "data pilot source has no canonical WorkbookOptions mapping",
      );
    }
    const source = parseSheetAddress(attributeString(sourceRange, "table:cell-range-address"));
    if (!source?.address) {
      throw unknownOdsElement(
        sourceRange,
        CONTENT_PATH,
        "table:cell-range-address",
        "data pilot source range is missing",
      );
    }
    const sourceWorksheet = source?.sheet
      ? worksheets.find((entry) => entry.name === source.sheet)
      : worksheet;
    if (!sourceWorksheet) {
      throw unknownOdsElement(
        sourceRange,
        CONTENT_PATH,
        source?.sheet ?? "table:cell-range-address",
        "data pilot source worksheet is missing",
      );
    }
    const grandTotal = attributeString(table, "table:grand-total");
    if (grandTotal !== undefined && !["none", "row", "column", "both"].includes(grandTotal)) {
      throw unknownOdsElement(
        table,
        CONTENT_PATH,
        "table:grand-total",
        `invalid data pilot grand-total value: ${grandTotal}`,
      );
    }
    const drillDown = parseOdsBoolean(table, "table:drill-down-on-double-click");
    const rows: string[] = [];
    const columns: string[] = [];
    const pages: PivotPageFieldOptions[] = [];
    const data: PivotDataField[] = [];
    const hidden: string[] = [];
    for (const field of childrenNamed(table, "table:data-pilot-field")) {
      const name = attributeString(field, "table:source-field-name") ?? "";
      const orientation = attributeString(field, "table:orientation");
      if (orientation === undefined) {
        throw unknownOdsElement(
          field,
          CONTENT_PATH,
          "table:orientation",
          "data pilot field orientation is required",
        );
      }
      if (
        attributeString(field, "table:is-data-layout-field") !== undefined ||
        attributeString(field, "table:used-hierarchy") !== undefined ||
        (field.elements ?? []).some((child) => child.type === "element")
      ) {
        throw unknownOdsElement(
          field,
          CONTENT_PATH,
          "table:data-pilot-field",
          "field hierarchy, grouping, and layout metadata have no canonical pivot mapping",
        );
      }
      if (orientation === "row") rows.push(name);
      else if (orientation === "page") {
        const selectedPage = attributeString(field, "table:selected-page");
        const selectedIndex =
          selectedPage === undefined
            ? undefined
            : pivotFieldMembers(sourceWorksheet, source.address, name).indexOf(selectedPage);
        if (selectedPage !== undefined && (selectedIndex === undefined || selectedIndex < 0)) {
          throw unknownOdsElement(
            field,
            CONTENT_PATH,
            "table:selected-page",
            "selected page member is not present in the source range",
          );
        }
        pages.push({
          field: name,
          ...(selectedIndex === undefined ? {} : { item: selectedIndex }),
        });
      } else if (orientation === "column") columns.push(name);
      else if (orientation === "hidden") hidden.push(name);
      else if (orientation === "data") {
        const functionValue = attributeString(field, "table:function");
        const summarize =
          functionValue === undefined || functionValue === "auto"
            ? undefined
            : ODF_PIVOT_FUNCTIONS[functionValue as keyof typeof ODF_PIVOT_FUNCTIONS];
        if (functionValue !== undefined && functionValue !== "auto" && !summarize) {
          const tableName = attributeString(table, "table:name");
          functions.push({
            field: name,
            summaryFunction: functionValue,
            ...(tableName ? { tableName } : {}),
          });
        }
        data.push({
          field: name,
          summarize,
        });
      }
    }
    worksheet.pivotTables = [
      ...(worksheet.pivotTables ?? []),
      {
        mode: "source",
        name: attributeString(table, "table:name"),
        source: source.address,
        sourceSheet: source?.sheet,
        location: target.address.split(":")[0],
        rows,
        ...(pages.length > 0 ? { pages } : {}),
        columns,
        data,
        rowGrandTotals:
          grandTotal === undefined ? undefined : grandTotal === "row" || grandTotal === "both",
        colGrandTotals:
          grandTotal === undefined ? undefined : grandTotal === "column" || grandTotal === "both",
        enableDrill: drillDown,
        ...(hidden.length > 0 ? { hiddenFields: hidden } : {}),
      },
    ];
  }
  return functions;
}

function parseOdsBoolean(element: Element, name: string): boolean | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  if (value === "true" || value === "1") return true;
  if (value === "false" || value === "0") return false;
  throw unknownOdsElement(element, CONTENT_PATH, name, `invalid boolean value: ${value}`);
}
