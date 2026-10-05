import {
  toUint8Array,
  type ChartSpaceOptions,
  type FormContainerOptions,
  type FormControlOptions,
} from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  CHART_MIME,
  chartBodyXml,
  childNamed,
  childrenNamed,
  escapeText,
  generateOcf,
  hasOcfManifestOverlay,
  metaXml,
  OdfSchemaError,
  parseEmbeddedCharts,
  parseMeta,
  readOcf,
  readXml,
  textOf,
  xmlElement,
  type OcfManifestOptions,
  type OdfPackageFiles,
} from "@office-open/odf";
import type { WorkbookOptions } from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { parseWorksheetCharts } from "./drawing-chart";
import { OdsParseError } from "./error";
import { parseNumberStyles } from "./numbering";
import type { OdsDocumentOptions } from "./semantics";
import type { OdsPackageMemberOptions } from "./semantics";
import { base64ToBytes, bytesToBase64, imageMediaType } from "./shared-data";
import { parseDimensionStyles, parseStyleOverlays, styleOverlaysXml } from "./styles";
import { worksheet, worksheetXml } from "./worksheet";

export const MIME = "application/vnd.oasis.opendocument.spreadsheet";

export const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"',
  'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"',
  'xmlns:xforms="http://www.w3.org/2002/xforms"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
].join(" ");

export function generateWorkbook(options: OdsDocumentOptions): Uint8Array {
  const { packageManifest, packageMembers, styleOverlays } = options;
  const styles: string[] = styleOverlaysXml(styleOverlays);
  const chartFrames = (options.worksheets ?? []).flatMap((worksheet, worksheetIndex) =>
    (worksheet.charts ?? []).map((chart, chartIndex) => ({
      name: chart.name ?? `Object ${worksheetIndex + chartIndex + 1}`,
      worksheet: worksheet.name ?? `Sheet${worksheetIndex + 1}`,
      chart,
    })),
  );
  const sheets = (options.worksheets ?? []).map((worksheet, index) =>
    worksheetXml(worksheet, index + 1, styles, chartFrames),
  );
  const files: OdfPackageFiles = {
    "content.xml": contentXml(
      sheets.join(""),
      styles,
      options.definedNames,
      options.calculation,
      options.forms,
    ),
    "styles.xml": stylesXml(),
    "meta.xml": metaXml(options),
  };
  for (const member of packageMembers ?? []) addPackageFile(files, member);
  for (const entry of chartFrames) files[`${entry.name}/content.xml`] = chartBodyXml(entry.chart);
  for (const worksheet of options.worksheets ?? []) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        for (const graphic of cell.graphics ?? []) {
          if (graphic.type === "image") {
            if (graphic.data !== undefined) files[graphic.href] = base64ToBytes(graphic.data);
            else if (graphic.sourceUrl === undefined)
              throw new OdsParseError(
                `content.xml: ${graphic.href}: image data or source URL is missing`,
                "content.xml",
                "/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/draw:frame/draw:image",
                "draw:image",
                "missing-image-source",
              );
          } else if (graphic.chart) {
            files[`${graphic.href}/content.xml`] = chartBodyXml(graphic.chart);
          }
        }
      }
    }
  }
  return generateOcf(
    MIME,
    files,
    Object.fromEntries([
      ...chartFrames.map((entry) => [`${entry.name}/`, CHART_MIME]),
      ...(options.worksheets ?? []).flatMap((worksheet) =>
        (worksheet.rows ?? []).flatMap((row) =>
          (row.cells ?? []).flatMap((cell) =>
            (cell.graphics ?? [])
              .filter(
                (graphic) =>
                  graphic.type === "image" &&
                  graphic.data !== undefined &&
                  graphic.sourceUrl === undefined,
              )
              .map((graphic) => [graphic.href, imageMediaType(graphic.href)]),
          ),
        ),
      ),
    ]),
    packageManifest,
  );
}

function addPackageFile(files: OdfPackageFiles, member: OdsPackageMemberOptions): void {
  if (
    !member.path ||
    member.path.startsWith("/") ||
    member.path.endsWith("/") ||
    member.path.split("/").includes("..")
  )
    throw new Error(`Invalid ODS package member path: ${member.path}`);
  if (member.path === "META-INF/manifest.xml")
    throw new Error("ODS package members cannot replace META-INF/manifest.xml");
  if (files[member.path] !== undefined)
    throw new Error(`ODS package member conflicts with modeled content: ${member.path}`);
  files[member.path] = typeof member.data === "string" ? member.data : toUint8Array(member.data);
}

export function unsupportedOdsValue(name: string, reason: string): OdsParseError {
  return new OdsParseError(
    `content.xml: ${name} ${reason}`,
    "content.xml",
    "/office:document-content/office:body/office:spreadsheet",
    name,
    "no canonical ODS mapping",
  );
}

function isExternalUrl(value: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("../");
}

export function parseWorkbook(data: Uint8Array): OdsDocumentOptions {
  try {
    return parseOdsBody(data);
  } catch (cause) {
    if (cause instanceof OdsParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdsParseError(cause.message, cause.part, cause.path, cause.name, cause.reason, {
        cause,
      });
    }
    throw new OdsParseError(
      cause instanceof Error ? cause.message : "Unable to parse ODS package",
      "mimetype",
      "/mimetype",
      "mimetype",
      "invalid ODS package",
      { cause },
    );
  }
}

function parseOdsBody(data: Uint8Array): OdsDocumentOptions {
  const { files, binaries, manifest } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:spreadsheet");
  const automaticStyles = childNamed(content, "office:automatic-styles");
  const dimensions = parseDimensionStyles(automaticStyles);
  const cellStyles = parseNumberStyles(automaticStyles);
  const chartPool = parseEmbeddedCharts(manifest, files);
  const usedBinaryPaths = new Set<string>();
  const calcProperties = parseOdsSemantics(body);
  const forms = parseForms(body);
  const embeddedCharts = [...parseWorksheetCharts(body, chartPool)];
  const definedNames = parseDefinedNames(body) ?? [];
  const worksheets = childrenNamed(body, "table:table").map((table, index) => {
    const parsed = worksheet(table, index + 1, dimensions, cellStyles);
    const name = attributeString(table, "table:name");
    const charts = embeddedCharts
      .filter((entry) => entry.worksheet === (name ?? `Sheet${index + 1}`))
      .map((entry) => entry.chart);
    return charts.length > 0
      ? { ...parsed, charts: [...(parsed.charts ?? []), ...charts] }
      : parsed;
  });
  for (const worksheet of worksheets) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        for (const graphic of cell.graphics ?? []) {
          if (graphic.type === "image") {
            if (isExternalUrl(graphic.href)) {
              graphic.sourceUrl = graphic.href;
              graphic.data = undefined;
              continue;
            }
            const binary = binaries[graphic.href];
            if (binary) usedBinaryPaths.add(graphic.href);
            graphic.data = binary ? bytesToBase64(binary) : "";
            if (!graphic.data)
              throw new OdsParseError(
                `content.xml: ${graphic.href}: package image is missing`,
                "content.xml",
                "/draw:frame/draw:image/@xlink:href",
                "draw:image",
                "missing-package-part",
              );
          } else {
            const chart = chartPool.get(graphic.href);
            if (chart) graphic.chart = chart;
          }
        }
      }
    }
  }
  rejectUnknownSpreadsheetChildren(body);
  validateEmbeddedObjectBodies(body, files, binaries, chartPool);
  const result: WorkbookOptions = {
    ...parseMeta(files),
    ...(definedNames.length > 0 ? { definedNames } : {}),
    ...(calcProperties ? { calculation: calcProperties } : {}),
    ...(forms.length > 0 ? { forms } : {}),
    worksheets,
  };
  const styleOverlays = parseStyleOverlays(automaticStyles);
  const packageMembers = sourcePackageMembers(
    manifest,
    files,
    binaries,
    chartPool,
    usedBinaryPaths,
  );
  return {
    ...result,
    ...(hasOcfManifestOverlay(manifest) ? { packageManifest: manifest } : {}),
    ...(packageMembers.length > 0 ? { packageMembers } : {}),
    ...(styleOverlays.length > 0 ? { styleOverlays } : {}),
  };
}

function validateEmbeddedObjectBodies(
  body: Element | undefined,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
): void {
  const references = descendantElements(body).flatMap((element) =>
    element.name === "draw:object-ole"
      ? [attributeString(element, "xlink:href")?.replace(/^\.\//, "").replace(/^\//, "")]
      : [],
  );
  for (const href of references) {
    if (href && (binaries[href] || files[`${href}/content.xml`] || chartPool.has(href))) continue;
    throw new OdsParseError(
      `content.xml: ${href ?? ""}: embedded object subdocument is missing`,
      "content.xml",
      "/draw:frame/draw:object-ole/@xlink:href",
      "draw:object-ole",
      "embedded-object-subdocument-missing",
    );
  }
}

function sourcePackageMembers(
  manifest: OcfManifestOptions,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
  usedBinaryPaths: Set<string>,
): OdsPackageMemberOptions[] {
  const mediaTypes = new Map(
    manifest.entries.map((entry) => [entry.fullPath, entry.mediaType] as const),
  );
  const modeledPaths = new Set([
    "content.xml",
    "styles.xml",
    "meta.xml",
    "META-INF/manifest.xml",
    ...[...chartPool.keys()].map((path) => `${path}/content.xml`),
    ...usedBinaryPaths,
  ]);
  return [
    ...Object.entries(files).map(([path, data]) => ({ path, data })),
    ...Object.entries(binaries).map(([path, data]) => ({ path, data })),
  ]
    .filter(({ path }) => !modeledPaths.has(path))
    .sort((left, right) => left.path.localeCompare(right.path))
    .map(({ path, data }) => ({ path, mediaType: mediaTypes.get(path), data }));
}

export function contentXml(
  sheets: string,
  styles: string[],
  definedNames: WorkbookOptions["definedNames"],
  calculation?: WorkbookOptions["calculation"],
  forms?: WorkbookOptions["forms"],
): string {
  const expressions = `${formsXml(forms)}${semanticsXml(calculation)}${definedNamesXml(
    definedNames,
  )}`;
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:spreadsheet>${sheets}${expressions}</office:spreadsheet></office:body></office:document-content>`;
}

export function parseOdsSemantics(body: Element | undefined): WorkbookOptions["calculation"] {
  const calculation = childNamed(body, "table:calculation-settings");
  if (!calculation) return undefined;
  const iteration = childNamed(calculation, "table:iteration");
  const result: NonNullable<WorkbookOptions["calculation"]> = {
    ...(iteration
      ? {
          iterate: attributeString(iteration, "table:status") === "enable",
          ...(attributeNumber(iteration, "table:steps") !== undefined
            ? { iterateCount: attributeNumber(iteration, "table:steps") }
            : {}),
          ...(attributeNumber(iteration, "table:maximum-difference") !== undefined
            ? { iterateDelta: attributeNumber(iteration, "table:maximum-difference") }
            : {}),
        }
      : {}),
    ...(attributeString(calculation, "table:case-sensitive") === "false"
      ? { caseSensitive: false }
      : {}),
    ...(attributeString(calculation, "table:precision-as-shown") === "true"
      ? { precisionAsShown: true }
      : {}),
    ...(attributeString(calculation, "table:automatic-find-labels") === "true"
      ? { automaticLabelSearch: true }
      : {}),
    ...(attributeString(calculation, "table:use-regular-expressions") === "true"
      ? { regularExpressionSearch: true }
      : {}),
    ...(attributeString(calculation, "table:use-wildcards") === "true"
      ? { wildcardSearch: true }
      : {}),
  };
  const nullDate = childNamed(calculation, "table:null-date");
  if (nullDate) {
    result.nullDate = {
      year: attributeNumber(nullDate, "table:date-value-year") ?? 1899,
      month: attributeNumber(nullDate, "table:date-value-month") ?? 12,
      day: attributeNumber(nullDate, "table:date-value-day") ?? 30,
    };
  }
  assertCalculationChildren(calculation, result);
  return result;
}

export function assertCalculationChildren(
  calculation: Element,
  result: NonNullable<WorkbookOptions["calculation"]>,
): void {
  for (const child of calculation.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "table:iteration") {
      assertElementNames(child, ["table:status", "table:steps", "table:maximum-difference"]);
      continue;
    }
    if (child.name === "table:null-date") {
      assertElementNames(child, [
        "table:date-value-year",
        "table:date-value-month",
        "table:date-value-day",
      ]);
      if (!result.nullDate) result.nullDate = { year: 1899, month: 12, day: 30 };
      continue;
    }
    throw unknownOdsElement(
      child,
      "/office:document-content/office:body/office:spreadsheet/table:calculation-settings",
      child.name ?? "",
      "element has no canonical calculation mapping",
    );
  }
}

export function formsXml(forms: WorkbookOptions["forms"]): string {
  if (!forms?.length) return "";
  return xmlElement(
    "office:forms",
    {
      "form:automatic-focus": forms[0]?.automaticFocus,
      "form:apply-design-mode": forms[0]?.designMode,
    },
    forms.map((form) =>
      xmlElement(
        "form:form",
        { "form:name": form.name },
        form.controls.map((control) => formControlXml(control)),
      ),
    ),
  );
}

export function formControlXml(control: FormControlOptions): string {
  const common = {
    "form:id": control.id,
    "form:name": control.name,
    "form:disabled": control.disabled,
    "form:tab-index": control.tabIndex,
    "form:automatic-focus": control.automaticFocus,
  };
  if (control.control === "checkBox")
    return xmlElement("form:checkbox", {
      ...common,
      "form:current-state": control.checked ? "checked" : "unchecked",
    });
  if (control.control === "dropDownList")
    return xmlElement(
      "form:listbox",
      common,
      control.entries.map((label, optionIndex) =>
        xmlElement(
          "form:option",
          { "form:selected": control.selectedIndex === optionIndex || undefined },
          [escapeText(label)],
        ),
      ),
    );
  return xmlElement("form:text", {
    ...common,
    "office:value-type": control.valueType,
    "form:current-value": control.value,
    "form:max-length": control.maxLength,
  });
}

export function parseForms(body: Element | undefined): FormContainerOptions[] {
  return descendantElements(body)
    .filter((element) => element.name === "office:forms")
    .flatMap((element) => {
      assertElementNames(
        element,
        ["form:automatic-focus", "form:apply-design-mode"],
        ["form:form"],
      );
      return childrenNamed(element, "form:form").map((form) => {
        assertElementNames(form, ["form:name"], ["form:text", "form:checkbox", "form:listbox"]);
        return {
          ...(attributeString(form, "form:name")
            ? { name: attributeString(form, "form:name") }
            : {}),
          ...(attributeString(element, "form:automatic-focus") === "true"
            ? { automaticFocus: true }
            : {}),
          ...(attributeString(element, "form:apply-design-mode") === "true"
            ? { designMode: true }
            : {}),
          controls:
            form.elements
              ?.filter((child) => child.type === "element")
              .map((control) => parseFormControl(control)) ?? [],
        };
      });
    });
}

export function descendantElements(element: Element | undefined): Element[] {
  return (element?.elements ?? []).flatMap((child) =>
    child.type === "element" ? [child, ...descendantElements(child)] : [],
  );
}

export function parseFormControl(control: Element): FormControlOptions {
  assertElementNames(
    control,
    [
      "form:id",
      "form:name",
      "form:disabled",
      "form:tab-index",
      "form:automatic-focus",
      "office:value-type",
      "form:current-value",
      "form:max-length",
      "form:current-state",
      "form:selected",
    ],
    ["form:option", "form:item"],
  );
  const common = {
    ...(attributeString(control, "form:id") ? { id: attributeString(control, "form:id") } : {}),
    ...(attributeString(control, "form:name")
      ? { name: attributeString(control, "form:name") }
      : {}),
    ...(attributeString(control, "form:disabled") === "true" ? { disabled: true } : {}),
    ...(attributeNumber(control, "form:tab-index") !== undefined
      ? { tabIndex: attributeNumber(control, "form:tab-index") }
      : {}),
    ...(attributeString(control, "form:automatic-focus") === "true"
      ? { automaticFocus: true }
      : {}),
  };
  if (control.name === "form:text")
    return {
      control: "text",
      ...common,
      ...(attributeString(control, "form:current-value")
        ? { value: attributeString(control, "form:current-value") }
        : {}),
      ...(attributeString(control, "office:value-type")
        ? { valueType: attributeString(control, "office:value-type") as "float" | "string" }
        : {}),
      ...(attributeNumber(control, "form:max-length") !== undefined
        ? { maxLength: attributeNumber(control, "form:max-length") }
        : {}),
    };
  if (control.name === "form:checkbox") {
    const state = attributeString(control, "form:current-state");
    if (state !== "checked" && state !== "unchecked")
      throw unknownOdsElement(control, "", control.name ?? "", "invalid checkbox state");
    return { control: "checkBox", ...common, checked: state === "checked" };
  }
  if (control.name === "form:listbox") {
    const entries = childrenNamed(control, "form:option").map((option) => textOf(option));
    const selected = childrenNamed(control, "form:option").findIndex(
      (option) => attributeString(option, "form:selected") === "true",
    );
    return {
      control: "dropDownList",
      ...common,
      entries,
      ...(selected >= 0 ? { selectedIndex: selected } : {}),
    };
  }
  throw unknownOdsElement(
    control,
    "/office:spreadsheet/office:forms",
    control.name ?? "",
    "element has no canonical control mapping",
  );
}

export function semanticsXml(calculation: WorkbookOptions["calculation"]): string {
  if (!calculation) return "";
  const hasSettings =
    calculation.iterate !== undefined ||
    calculation.caseSensitive !== undefined ||
    calculation.precisionAsShown !== undefined ||
    calculation.automaticLabelSearch !== undefined ||
    calculation.regularExpressionSearch !== undefined ||
    calculation.wildcardSearch !== undefined ||
    calculation.nullDate !== undefined;
  if (!hasSettings) return "";
  const iteration = {
    "table:status":
      calculation.iterate === undefined ? undefined : calculation.iterate ? "enable" : "disable",
    "table:steps": calculation.iterateCount,
    "table:maximum-difference": calculation.iterateDelta,
  };
  return xmlElement(
    "table:calculation-settings",
    {
      "table:case-sensitive": calculation.caseSensitive,
      "table:precision-as-shown": calculation.precisionAsShown,
      "table:automatic-find-labels": calculation.automaticLabelSearch,
      "table:use-regular-expressions": calculation.regularExpressionSearch,
      "table:use-wildcards": calculation.wildcardSearch,
    },
    [
      ...(calculation.iterate !== undefined ? [xmlElement("table:iteration", iteration)] : []),
      ...(calculation.nullDate
        ? [
            xmlElement("table:null-date", {
              "table:date-value-year": calculation.nullDate.year,
              "table:date-value-month": calculation.nullDate.month,
              "table:date-value-day": calculation.nullDate.day,
            }),
          ]
        : []),
    ],
  );
}

export function definedNamesXml(definedNames: WorkbookOptions["definedNames"]): string {
  if (!definedNames?.length) return "";
  return xmlElement(
    "table:named-expressions",
    undefined,
    definedNames.map((definedName) =>
      /^[A-Za-z0-9_]+![A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$/.test(definedName.value)
        ? xmlElement("table:named-range", {
            "table:name": definedName.name,
            "table:cell-range-address": definedName.value,
          })
        : xmlElement("table:named-expression", {
            "table:name": definedName.name,
            "table:expression": definedName.value,
          }),
    ),
  );
}

export function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/></office:document-styles>`;
}

export function rejectUnknownSpreadsheetChildren(body: Element | undefined): void {
  const allowed = new Set([
    "table:table",
    "table:calculation-settings",
    "office:forms",
    "table:named-expressions",
    "draw:frame",
  ]);
  for (const child of body?.elements ?? []) {
    if (child.name && !allowed.has(child.name)) {
      const name = child.name;
      throw new OdsParseError(
        `content.xml: /office:document-content/office:body/office:spreadsheet/${name}: no canonical WorkbookOptions mapping`,
        "content.xml",
        `/office:document-content/office:body/office:spreadsheet/${name}`,
        name,
        "no canonical WorkbookOptions mapping",
      );
    }
  }
}

export function parseDefinedNames(body: Element | undefined): WorkbookOptions["definedNames"] {
  const expressions = childNamed(body, "table:named-expressions");
  return [
    ...childrenNamed(expressions, "table:named-range").map((range) => ({
      name: attributeString(range, "table:name") ?? "",
      value: attributeString(range, "table:cell-range-address") ?? "",
    })),
    ...childrenNamed(expressions, "table:named-expression").map((expression) => ({
      name: attributeString(expression, "table:name") ?? "",
      value: attributeString(expression, "table:expression") ?? "",
    })),
  ];
}

export function unknownOdsElement(
  element: Element,
  parent: string,
  name: string,
  reason: string,
): OdsParseError {
  const path = `${parent}/${element.name ?? name}`;
  return new OdsParseError(
    `content.xml: ${path}: ${name}: ${reason}`,
    "content.xml",
    path,
    name,
    reason,
  );
}

export function assertElementNames(
  element: Element,
  allowedAttributes: string[],
  allowedChildren: string[] = [],
): void {
  const unknown = Object.keys(element.attributes ?? {}).filter(
    (name) => !allowedAttributes.includes(name),
  );
  if (unknown.length) {
    throw unknownOdsElement(
      element,
      `/office:document-content/office:body${element.name === "form:form" ? "/office:spreadsheet/office:forms" : ""}`,
      unknown[0] ?? "",
      "attribute has no canonical mapping",
    );
  }
  for (const child of element.elements ?? []) {
    if (child.type === "element" && !allowedChildren.includes(child.name ?? "")) {
      throw unknownOdsElement(
        element,
        `/office:document-content/office:body${element.name === "form:form" ? "/office:spreadsheet/office:forms" : ""}`,
        child.name ?? "",
        "child has no canonical mapping",
      );
    }
  }
}
