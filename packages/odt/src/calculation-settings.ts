import { attributeNumber, attributeString, childNamed, xmlElement } from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdtParseError } from "./error";

/** Iteration controls for circular formula resolution in ODF table declarations. */
export interface OdtIterationOptions {
  enabled?: boolean;
  steps?: number;
  maximumDifference?: number;
}

/** Date value used as day zero for ODF serial dates. */
export interface OdtNullDateOptions {
  year: number;
  month: number;
  day: number;
}

/** ODF table declarations shared by every office content family. */
export interface OdtCalculationSettingsOptions {
  caseSensitive?: boolean;
  precisionAsShown?: boolean;
  searchCriteriaMustApplyToWholeCell?: boolean;
  automaticFindLabels?: boolean;
  regularExpressionSearch?: boolean;
  wildcardSearch?: boolean;
  nullYear?: number;
  nullDate?: OdtNullDateOptions;
  iteration?: OdtIterationOptions;
}

const BOOLEAN_VALUES = new Map([
  ["true", true],
  ["false", false],
]);

function booleanAttribute(element: Element, name: string): boolean | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  const parsed = BOOLEAN_VALUES.get(value);
  if (parsed === undefined)
    throw new OdtParseError(
      `content.xml: invalid ${name} boolean: ${value}`,
      "content.xml",
      `/${element.name}/${name}`,
      name,
      "invalid-boolean",
    );
  return parsed;
}

function assertKnownAttributes(element: Element, names: string[]): void {
  const unknown = Object.keys(element.attributes ?? {}).filter((name) => !names.includes(name));
  if (unknown.length > 0)
    throw new OdtParseError(
      `content.xml: unknown ${element.name} attribute: ${unknown.join(", ")}`,
      "content.xml",
      `/${element.name}`,
      unknown[0]!,
      "attribute has no canonical mapping",
    );
}

/** Parse the table declaration prelude shared by ODF content documents. */
export function parseCalculationSettings(element: Element): OdtCalculationSettingsOptions {
  assertKnownAttributes(element, [
    "table:case-sensitive",
    "table:precision-as-shown",
    "table:search-criteria-must-apply-to-whole-cell",
    "table:automatic-find-labels",
    "table:use-regular-expressions",
    "table:use-wildcards",
    "table:null-year",
  ]);
  const result: OdtCalculationSettingsOptions = {
    ...(booleanAttribute(element, "table:case-sensitive") !== undefined
      ? { caseSensitive: booleanAttribute(element, "table:case-sensitive") }
      : {}),
    ...(booleanAttribute(element, "table:precision-as-shown") !== undefined
      ? { precisionAsShown: booleanAttribute(element, "table:precision-as-shown") }
      : {}),
    ...(booleanAttribute(element, "table:search-criteria-must-apply-to-whole-cell") !== undefined
      ? {
          searchCriteriaMustApplyToWholeCell: booleanAttribute(
            element,
            "table:search-criteria-must-apply-to-whole-cell",
          ),
        }
      : {}),
    ...(booleanAttribute(element, "table:automatic-find-labels") !== undefined
      ? { automaticFindLabels: booleanAttribute(element, "table:automatic-find-labels") }
      : {}),
    ...(booleanAttribute(element, "table:use-regular-expressions") !== undefined
      ? { regularExpressionSearch: booleanAttribute(element, "table:use-regular-expressions") }
      : {}),
    ...(booleanAttribute(element, "table:use-wildcards") !== undefined
      ? { wildcardSearch: booleanAttribute(element, "table:use-wildcards") }
      : {}),
    ...(attributeNumber(element, "table:null-year") !== undefined
      ? { nullYear: attributeNumber(element, "table:null-year") }
      : {}),
  };
  const nullDate = childNamed(element, "table:null-date");
  if (nullDate) {
    assertKnownAttributes(nullDate, ["table:value-type", "table:date-value"]);
    const isoDate = attributeString(nullDate, "table:date-value");
    const match = /^(\d{4})-(\d{2})-(\d{2})/u.exec(isoDate ?? "");
    if (isoDate !== undefined && !match)
      throw new OdtParseError(
        `content.xml: unsupported table:null-date value: ${isoDate}`,
        "content.xml",
        "/table:calculation-settings/table:null-date",
        "table:date-value",
        "unsupported-date-value",
      );
    result.nullDate = match
      ? { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) }
      : { year: 1899, month: 12, day: 30 };
  }
  const iteration = childNamed(element, "table:iteration");
  if (iteration) {
    assertKnownAttributes(iteration, ["table:status", "table:steps", "table:maximum-difference"]);
    const status = attributeString(iteration, "table:status");
    if (status !== undefined && status !== "enable" && status !== "disable")
      throw new OdtParseError(
        `content.xml: invalid table:iteration status: ${status}`,
        "content.xml",
        "/table:calculation-settings/table:iteration",
        "table:status",
        "invalid-iteration-status",
      );
    result.iteration = {
      ...(status ? { enabled: status === "enable" } : {}),
      ...(attributeNumber(iteration, "table:steps") !== undefined
        ? { steps: attributeNumber(iteration, "table:steps") }
        : {}),
      ...(attributeNumber(iteration, "table:maximum-difference") !== undefined
        ? { maximumDifference: attributeNumber(iteration, "table:maximum-difference") }
        : {}),
    };
  }
  return result;
}

/** Serialize the table declaration prelude shared by ODF content documents. */
export function calculationSettingsXml(options: OdtCalculationSettingsOptions): string {
  return xmlElement(
    "table:calculation-settings",
    {
      "table:case-sensitive": options.caseSensitive,
      "table:precision-as-shown": options.precisionAsShown,
      "table:search-criteria-must-apply-to-whole-cell": options.searchCriteriaMustApplyToWholeCell,
      "table:automatic-find-labels": options.automaticFindLabels,
      "table:use-regular-expressions": options.regularExpressionSearch,
      "table:use-wildcards": options.wildcardSearch,
      "table:null-year": options.nullYear,
    },
    [
      ...(options.nullDate
        ? [
            xmlElement("table:null-date", {
              "table:value-type": "date",
              "table:date-value": `${String(options.nullDate.year).padStart(4, "0")}-${String(
                options.nullDate.month,
              ).padStart(2, "0")}-${String(options.nullDate.day).padStart(2, "0")}`,
            }),
          ]
        : []),
      ...(options.iteration
        ? [
            xmlElement("table:iteration", {
              "table:status":
                options.iteration.enabled === undefined
                  ? undefined
                  : options.iteration.enabled
                    ? "enable"
                    : "disable",
              "table:steps": options.iteration.steps,
              "table:maximum-difference": options.iteration.maximumDifference,
            }),
          ]
        : []),
    ],
  );
}
