import { attributeNumber, attributeString, xmlElement } from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { unknownOdsElement } from "./workbook";

type ColumnTransform = {
  columns: number[];
};

export type DataMappingTransformationOptions =
  | ({ kind: "removeColumns" } & ColumnTransform)
  | { kind: "splitColumn"; column: number; separator: string }
  | ({ kind: "mergeColumns"; mergeString: string } & ColumnTransform)
  | { kind: "sortColumns"; fieldNumber: number; ascending?: boolean }
  | ({ kind: "text"; type: "lowercase" | "uppercase" | "capitalize" | "trim" } & ColumnTransform)
  | ({
      kind: "aggregate";
      type: "sum" | "average" | "min" | "max";
    } & ColumnTransform)
  | ({
      kind: "number";
      type:
        | "round"
        | "roundUp"
        | "roundDown"
        | "absolute"
        | "logE"
        | "log10"
        | "cube"
        | "square"
        | "squareRoot"
        | "exponential"
        | "even"
        | "odd"
        | "sign";
      precision: number;
    } & ColumnTransform)
  | ({ kind: "replaceNull"; replaceString: string } & ColumnTransform)
  | ({
      kind: "dateTime";
      type:
        | "dateString"
        | "year"
        | "startOfYear"
        | "endOfYear"
        | "month"
        | "monthName"
        | "startOfMonth"
        | "endOfMonth"
        | "day"
        | "dayOfWeek"
        | "dayOfYear"
        | "quarter"
        | "startOfQuarter"
        | "endOfQuarter"
        | "time"
        | "hour"
        | "minute"
        | "second";
    } & ColumnTransform);

export interface DataMappingOptions {
  href: string;
  provider: string;
  frequency?: number;
  id?: string;
  databaseName?: string;
  transformations: DataMappingTransformationOptions[];
}

const TRANSFORMATION_ELEMENTS: Record<string, DataMappingTransformationOptions["kind"]> = {
  "calcext:column-remove-transformation": "removeColumns",
  "calcext:column-split-transformation": "splitColumn",
  "calcext:column-merge-transformation": "mergeColumns",
  "calcext:column-sort-transformation": "sortColumns",
  "calcext:column-text-transformation": "text",
  "calcext:column-aggregate-transformation": "aggregate",
  "calcext:column-number-transformation": "number",
  "calcext:column-replacenull-transformation": "replaceNull",
  "calcext:column-datetime-transformation": "dateTime",
};

export function parseDataMappings(body: Element | undefined): DataMappingOptions[] {
  const container = elementChild(body, "calcext:data-mappings");
  if (!container) return [];
  assertNames(container, ["calcext:data-mapping"]);
  return elementChildren(container, "calcext:data-mapping").map(parseMapping);
}

export function dataMappingsXml(mappings: readonly DataMappingOptions[]): string {
  if (!mappings.length) return "";
  return xmlElement(
    "calcext:data-mappings",
    undefined,
    mappings.map((mapping) =>
      xmlElement(
        "calcext:data-mapping",
        {
          "xlink:href": mapping.href,
          "calcext:provider": mapping.provider,
          "calcext:data-frequency": mapping.frequency,
          "calcext:id": mapping.id,
          "calcext:database-name": mapping.databaseName,
        },
        [
          xmlElement(
            "calcext:data-transformations",
            undefined,
            mapping.transformations.map(transformationXml),
          ),
        ],
      ),
    ),
  );
}

function parseMapping(mapping: Element): DataMappingOptions {
  assertNames(mapping, ["calcext:data-transformations"]);
  const transformations = elementChild(mapping, "calcext:data-transformations");
  assertNames(transformations, Object.keys(TRANSFORMATION_ELEMENTS));
  return {
    href: required(mapping, "xlink:href"),
    provider: required(mapping, "calcext:provider"),
    ...(attributeNumber(mapping, "calcext:data-frequency") === undefined
      ? {}
      : { frequency: requiredNumber(mapping, "calcext:data-frequency") }),
    id: attributeString(mapping, "calcext:id"),
    databaseName: attributeString(mapping, "calcext:database-name"),
    transformations: elementChildren(transformations).map(parseTransformation),
  };
}

function requiredNumber(element: Element, name: string): number {
  const value = attributeNumber(element, name);
  if (value === undefined || !Number.isSafeInteger(value) || value < 0) {
    throw unknownOdsElement(
      element,
      "/office:document-content/office:body/office:spreadsheet/calcext:data-mappings",
      name,
      "numeric property is missing or invalid",
    );
  }
  return value;
}

function parseTransformation(element: Element): DataMappingTransformationOptions {
  const kind = TRANSFORMATION_ELEMENTS[element.name ?? ""];
  if (!kind) {
    throw unknownOdsElement(
      element,
      "/office:document-content/office:body/office:spreadsheet/calcext:data-mappings",
      element.name ?? "",
      "element has no canonical DataMappingOptions mapping",
    );
  }
  const columns = elementChildren(element, "calcext:column").map((column) =>
    requiredNumber(column, "calcext:column"),
  );
  if (kind === "removeColumns")
    return {
      kind,
      columns,
    };
  if (kind === "splitColumn")
    return {
      kind,
      column: requiredNumber(element, "calcext:column"),
      separator: required(element, "calcext:separator"),
    };
  if (kind === "mergeColumns")
    return { kind, columns, mergeString: required(element, "calcext:merge-string") };
  if (kind === "sortColumns") {
    const sort = elementChild(element, "table:sort");
    const sortBy = sort ? elementChildren(sort, "table:sort-by")[0] : undefined;
    if (!sortBy) {
      throw unknownOdsElement(
        element,
        "/office:document-content/office:body/office:spreadsheet/calcext:data-mappings",
        "table:sort-by",
        "sort transformation is missing its key",
      );
    }
    return {
      kind,
      fieldNumber: requiredNumber(sortBy, "table:field-number"),
      ...(attributeString(sortBy, "table:order") === "descending" ? { ascending: false } : {}),
    };
  }
  if (kind === "number")
    return {
      kind,
      columns,
      type: mapValues(element, NUMBER_TYPES),
      precision: requiredNumber(element, "calcext:precision"),
    };
  if (kind === "replaceNull")
    return {
      kind,
      columns,
      replaceString: required(element, "calcext:replace-string"),
    };
  const type = required(element, "calcext:type");
  if (kind === "text") return { kind, columns, type: mapValue(element, TEXT_TYPES, type) };
  if (kind === "aggregate")
    return { kind, columns, type: mapValue(element, AGGREGATE_TYPES, type) };
  return { kind, columns, type: mapValue(element, DATE_TIME_TYPES, type) };
}

function transformationXml(transformation: DataMappingTransformationOptions): string {
  const names: Record<DataMappingTransformationOptions["kind"], string> = {
    removeColumns: "calcext:column-remove-transformation",
    splitColumn: "calcext:column-split-transformation",
    mergeColumns: "calcext:column-merge-transformation",
    sortColumns: "calcext:column-sort-transformation",
    text: "calcext:column-text-transformation",
    aggregate: "calcext:column-aggregate-transformation",
    number: "calcext:column-number-transformation",
    replaceNull: "calcext:column-replacenull-transformation",
    dateTime: "calcext:column-datetime-transformation",
  };
  const children =
    "columns" in transformation
      ? transformation.columns.map((column) =>
          xmlElement("calcext:column", { "calcext:column": column }),
        )
      : [];
  if (transformation.kind === "sortColumns") {
    return xmlElement(names.sortColumns, undefined, [
      xmlElement("table:sort", undefined, [
        xmlElement("table:sort-by", {
          "table:field-number": transformation.fieldNumber,
          "table:order": transformation.ascending === false ? "descending" : "ascending",
        }),
      ]),
    ]);
  }
  const attributes: Record<string, string | number | boolean | undefined> = {};
  if (transformation.kind === "splitColumn") {
    attributes["calcext:column"] = transformation.column;
    attributes["calcext:separator"] = transformation.separator;
  }
  if (transformation.kind === "mergeColumns")
    attributes["calcext:merge-string"] = transformation.mergeString;
  if (transformation.kind === "number") {
    attributes["calcext:type"] = inverse(NUMBER_TYPES)[transformation.type];
    attributes["calcext:precision"] = transformation.precision;
  }
  if (transformation.kind === "replaceNull")
    attributes["calcext:replace-string"] = transformation.replaceString;
  if (transformation.kind === "text")
    attributes["calcext:type"] = inverse(TEXT_TYPES)[transformation.type];
  if (transformation.kind === "aggregate")
    attributes["calcext:type"] = inverse(AGGREGATE_TYPES)[transformation.type];
  if (transformation.kind === "dateTime")
    attributes["calcext:type"] = inverse(DATE_TIME_TYPES)[transformation.type];
  return xmlElement(names[transformation.kind], attributes, children);
}

const TEXT_TYPES = {
  lowercase: "lowercase",
  uppercase: "uppercase",
  capitalize: "capitalize",
  trim: "trim",
};
const AGGREGATE_TYPES = {
  sum: "sum",
  average: "average",
  min: "min",
  max: "max",
};
const NUMBER_TYPES = {
  round: "round",
  roundUp: "round-up",
  roundDown: "round-down",
  absolute: "abs",
  logE: "log",
  log10: "log10",
  cube: "cube",
  square: "square",
  squareRoot: "square-root",
  exponential: "exponential",
  even: "even",
  odd: "odd",
  sign: "sign",
};
const DATE_TIME_TYPES = {
  dateString: "date-string",
  year: "year",
  startOfYear: "start-of-year",
  endOfYear: "end-of-year",
  month: "month",
  monthName: "month-name",
  startOfMonth: "start-of-month",
  endOfMonth: "end-of-month",
  day: "day",
  dayOfWeek: "day-of-week",
  dayOfYear: "day-of-year",
  quarter: "quarter",
  startOfQuarter: "start-of-quarter",
  endOfQuarter: "end-of-quarter",
  time: "time",
  hour: "hour",
  minute: "minute",
  second: "seconds",
};

function elementChild(element: Element | undefined, name: string | undefined): Element | undefined {
  return name === undefined
    ? undefined
    : (element?.elements?.find((child) => child.name === name) ?? undefined);
}

function elementChildren(element: Element | undefined, name?: string): Element[] {
  return (element?.elements ?? []).filter((child) => name === undefined || child.name === name);
}

function assertNames(element: Element | undefined, allowed: readonly string[]): void {
  for (const child of element?.elements ?? []) {
    if (!allowed.includes(child.name ?? "")) {
      throw unknownOdsElement(
        child,
        "/office:document-content/office:body/office:spreadsheet/calcext:data-mappings",
        child.name ?? "",
        "child has no canonical DataMappingOptions mapping",
      );
    }
  }
}

function required(element: Element, name: string): string {
  const value = attributeString(element, name);
  if (value === undefined) {
    throw unknownOdsElement(
      element,
      "/office:document-content/office:body/office:spreadsheet/calcext:data-mappings",
      name,
      "required attribute is missing",
    );
  }
  return value;
}

function mapValues<Type extends string>(
  element: Element,
  values: Readonly<Record<Type, string>>,
): Type {
  return mapValue(element, values, required(element, "calcext:type"));
}

function mapValue<Type extends string>(
  element: Element,
  values: Readonly<Record<Type, string>>,
  token: string,
): Type {
  const entry = Object.entries(values).find(([, value]) => value === token);
  if (entry === undefined) {
    throw unknownOdsElement(
      element,
      "/office:document-content/office:body/office:spreadsheet/calcext:data-mappings",
      token,
      "value has no canonical DataMappingOptions mapping",
    );
  }
  return entry[0] as Type;
}

function inverse<T extends Record<string, string>>(values: T): Record<string, keyof T & string> {
  return Object.fromEntries(Object.entries(values).map(([key, value]) => [value, key]));
}
