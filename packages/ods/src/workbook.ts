import {
  convertOutput,
  toUint8Array,
  toUint8ArrayAsync,
  streamUint8Array,
  type DataType,
  type OutputByType,
  type OutputType,
  type SyncDataType,
} from "@office-open/core";
import type { ChartSpaceOptions, OfficeFormsOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  CHART_MIME,
  chartSubdocumentFiles,
  childNamed,
  childrenNamed,
  descendantElements,
  generateOcf,
  hasOcfManifestOptions,
  isExternalOdfUrl,
  metaXml,
  OdfSchemaError,
  parseEmbeddedCharts,
  officeFormsXml,
  parseOfficeForms,
  parseMeta,
  readOcfSync,
  readXml,
  unknownAttributeNames,
  unknownChildElement,
  xmlElement,
  type OcfManifestOptions,
  type OdfPackageFiles,
} from "@office-open/odf";
import type { PackageMemberOptions } from "@office-open/odf";
import type { DxfOptions, WorkbookOptions } from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { dataMappingsXml, parseDataMappings } from "./data-mappings";
import { parseWorksheetCharts } from "./drawing-chart";
import { OdsParseError } from "./error";
import { base64ToBytes, bytesToBase64, imageMediaType } from "./image-data";
import { parseNumberStyles } from "./numbering";
import type { OdsWorkbookOptions } from "./options";
import {
  attachContentValidations,
  contentValidationsXml,
  contentValidationNames,
  consolidationXml,
  databaseRangesXml,
  dataPilotTablesXml,
  namedExpressionsXml,
  parseCommonCellStyles,
  parseAutomaticCellStyleNames,
  parseConditionalFormats,
  parseContentValidations,
  parseConsolidation,
  parseDataPilotTables,
  parseDatabaseRanges,
  parseTableProtection,
  parseWorksheetGraphicStyles,
  parseWorksheetNamedExpressions,
} from "./spreadsheet";
import { parseDimensionStyles, parseStyleResources, styleResourcesXml } from "./styles";
import { worksheet, worksheetXml } from "./worksheet";

export const MIME = "application/vnd.oasis.opendocument.spreadsheet";
export const TEMPLATE_MIME = "application/vnd.oasis.opendocument.spreadsheet-template";
const LEGACY_MIME = "application/vnd.sun.xml.calc";

export const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"',
  'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"',
  'xmlns:xforms="http://www.w3.org/2002/xforms"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
  'xmlns:tableooo="http://openoffice.org/2009/table"',
  'xmlns:gnm="http://www.gnumeric.org/odf-extension/1.0"',
].join(" ");

function generateOdsPackage(options: OdsWorkbookOptions): Uint8Array {
  const { packageManifest, packageMembers, styleResources, packageMediaType } = options;
  const mimeType = packageMediaType ?? MIME;
  const styles: string[] = styleResourcesXml(styleResources);
  const chartFrames = (options.worksheets ?? []).flatMap((worksheet, worksheetIndex) =>
    (worksheet.charts ?? []).map((chart, chartIndex) => ({
      name: chart.name ?? `Object ${worksheetIndex + chartIndex + 1}`,
      worksheet: worksheet.name ?? `Sheet${worksheetIndex + 1}`,
      chart,
    })),
  );
  const validationNames = contentValidationNames(options.worksheets ?? []);
  const sheets = (options.worksheets ?? []).map((worksheet, index) =>
    worksheetXml(
      worksheet,
      index + 1,
      styles,
      chartFrames,
      namedExpressionsXml(
        (options.definedNames ?? []).filter((definedName) => definedName.localSheetId === index),
      ),
      validationNames,
      options.dxfs ?? [],
      index === 0 ? formsXml(options.forms) : "",
    ),
  );
  const files: OdfPackageFiles = {
    "content.xml": contentXml(sheets, styles, options),
    "styles.xml": stylesXml(),
    "meta.xml": metaXml(options),
  };
  for (const member of packageMembers ?? []) addPackageFile(files, member);
  for (const entry of chartFrames)
    Object.assign(files, chartSubdocumentFiles(`${entry.name}/`, entry.chart));
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
            Object.assign(files, chartSubdocumentFiles(`${graphic.href}/`, graphic.chart));
          }
        }
      }
    }
  }
  return generateOcf(
    mimeType,
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

function addPackageFile(files: OdfPackageFiles, member: PackageMemberOptions): void {
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

export function generateWorkbookSync<T extends OutputType = "uint8array">(
  options: OdsWorkbookOptions,
  output?: T,
): OutputByType[T] {
  return convertOutput(generateOdsPackage(options), (output ?? "uint8array") as T, MIME);
}

export async function generateWorkbook<T extends OutputType = "uint8array">(
  options: OdsWorkbookOptions,
  output?: T,
): Promise<OutputByType[T]> {
  return generateWorkbookSync(options, output);
}

export function generateWorkbookStream(options: OdsWorkbookOptions): ReadableStream<Uint8Array> {
  return streamUint8Array(generateOdsPackage(options));
}

export async function parseWorkbook(data: DataType): Promise<OdsWorkbookOptions> {
  return parseWorkbookSync(await toUint8ArrayAsync(data));
}

export function parseWorkbookSync(data: SyncDataType): OdsWorkbookOptions {
  try {
    return parseOdsPackage(toUint8Array(data));
  } catch (cause) {
    if (cause instanceof OdsParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdsParseError(
        cause.message,
        cause.part,
        cause.path,
        cause.elementName,
        cause.reason,
        { cause },
      );
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

function parseOdsPackage(data: Uint8Array): OdsWorkbookOptions {
  const { files, binaries, manifest, mimeType } = readOcfSync(data, [
    MIME,
    TEMPLATE_MIME,
    LEGACY_MIME,
  ]);
  const content = readXml(files, "content.xml");
  const stylesDocument =
    files["styles.xml"] === undefined ? undefined : readXml(files, "styles.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:spreadsheet");
  const automaticStyles = childNamed(content, "office:automatic-styles");
  const dimensions = parseDimensionStyles(automaticStyles);
  const cellStyles = parseNumberStyles(automaticStyles);
  const chartPool = parseEmbeddedCharts(manifest, files, binaries);
  const usedBinaryPaths = new Set<string>();
  const calcProperties = parseOdsSemantics(body);
  const forms = parseForms(body);
  const embeddedCharts = [...parseWorksheetCharts(body, chartPool)];
  const definedNames = parseDefinedNames(body) ?? [];
  const contentValidations = parseContentValidations(body);
  const dataMappings = parseDataMappings(body);
  const commonCellStyles = parseCommonCellStyles(stylesDocument);
  const graphicStyles = parseWorksheetGraphicStyles(automaticStyles, stylesDocument);
  const dxfs: DxfOptions[] = [];
  const localDefinedNames: NonNullable<WorkbookOptions["definedNames"]> = [];
  const worksheets = childrenNamed(body, "table:table").map((table, index) => {
    const parsed = worksheet(table, index + 1, dimensions, cellStyles, graphicStyles);
    const name = attributeString(table, "table:name");
    const charts = embeddedCharts
      .filter((entry) => entry.worksheet === (name ?? `Sheet${index + 1}`))
      .map((entry) => entry.chart);
    const named = parseWorksheetNamedExpressions(table, index);
    localDefinedNames.push(...named);
    parseTableProtection(table, parsed);
    const conditionalFormats = parseConditionalFormats(
      table,
      name ?? `Sheet${index + 1}`,
      new Map([
        ...cellStyles,
        ...commonCellStyles,
        ...parseAutomaticCellStyleNames(automaticStyles),
      ]),
      dxfs,
    );
    attachContentValidations(table, parsed, contentValidations);
    return {
      ...parsed,
      charts: charts.length > 0 ? [...(parsed.charts ?? []), ...charts] : parsed.charts,
      conditionalFormats: conditionalFormats.length ? conditionalFormats : undefined,
    };
  });
  parseDatabaseRanges(body, worksheets);
  parseDataPilotTables(body, worksheets);
  parseConsolidation(body, worksheets);
  definedNames.push(...localDefinedNames);
  for (const worksheet of worksheets) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        for (const graphic of cell.graphics ?? []) {
          if (graphic.type === "image") {
            if (isExternalOdfUrl(graphic.href)) {
              graphic.sourceUrl = graphic.href;
              graphic.data = undefined;
              continue;
            }
            const binary = binaries[graphic.href];
            if (binary) usedBinaryPaths.add(graphic.href);
            graphic.data = binary ? bytesToBase64(binary) : undefined;
            if (!graphic.data) {
              throw new OdsParseError(
                `content.xml: ${graphic.href}: image data is missing`,
                "content.xml",
                "/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/draw:frame/draw:image",
                "draw:image",
                "missing-image-data",
              );
            }
          } else {
            const chart = chartPool.get(graphic.href);
            if (chart) graphic.chart = chart;
          }
        }
      }
    }
  }
  rejectUnknownSpreadsheetChildren(body, dataMappings.length > 0);
  validateEmbeddedObjectBodies(body, files, binaries, chartPool);
  const result: WorkbookOptions = {
    ...parseMeta(files),
    ...(definedNames.length > 0 ? { definedNames } : {}),
    ...(calcProperties ? { calculation: calcProperties } : {}),
    ...(forms !== undefined ? { forms } : {}),
    ...(dxfs.length > 0 ? { dxfs } : {}),
    worksheets,
  };
  const styleResources = parseStyleResources(automaticStyles);
  const packageMembers = sourcePackageMembers(
    manifest,
    files,
    binaries,
    chartPool,
    usedBinaryPaths,
  );
  return {
    ...result,
    ...(dataMappings.length ? { dataMappings } : {}),
    ...(mimeType && mimeType !== MIME ? { packageMediaType: mimeType } : {}),
    ...(hasOcfManifestOptions(manifest) ? { packageManifest: manifest } : {}),
    ...(packageMembers.length > 0 ? { packageMembers } : {}),
    ...(styleResources.length > 0 ? { styleResources } : {}),
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
): PackageMemberOptions[] {
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
  sheets: readonly string[],
  styles: readonly string[],
  options: OdsWorkbookOptions,
): Uint8Array {
  const workbookDefinedNames = (options.definedNames ?? []).filter(
    (definedName) => definedName.localSheetId === undefined,
  );
  const declarations = contentValidationsXml(options.worksheets ?? []);
  const prelude = `${declarations}${semanticsXml(options.calculation)}`;
  const epilogue = `${definedNamesXml(workbookDefinedNames)}${dataMappingsXml(
    options.dataMappings ?? [],
  )}${databaseRangesXml(
    options.worksheets ?? [],
  )}${dataPilotTablesXml(options.worksheets ?? [])}${consolidationXml(options.worksheets ?? [])}`;
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const write = (value: string): void => {
    chunks.push(encoder.encode(value));
  };
  write(
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>`,
  );
  for (const style of styles) write(style);
  write("</office:automatic-styles><office:body><office:spreadsheet>");
  write(prelude);
  for (const sheet of sheets) write(sheet);
  write(epilogue);
  write("</office:spreadsheet></office:body></office:document-content>");
  const result = new Uint8Array(chunks.reduce((length, chunk) => length + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
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
    const isoDate = attributeString(nullDate, "table:date-value");
    result.nullDate = isoDate
      ? {
          year: Number(isoDate.slice(0, 4)),
          month: Number(isoDate.slice(5, 7)),
          day: Number(isoDate.slice(8, 10)),
        }
      : {
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
        "table:value-type",
        "table:date-value",
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
  return officeFormsXml(forms);
}

export function parseForms(body: Element | undefined): OfficeFormsOptions | undefined {
  if (body && childNamed(body, "office:forms")) {
    throw new OdsParseError(
      "content.xml: office:forms must be contained by table:table in a spreadsheet",
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/office:forms",
      "office:forms",
      "form container has no canonical WorkbookOptions placement",
    );
  }
  const tables = (body?.elements ?? []).filter(
    (child): child is Element => child.type === "element" && child.name === "table:table",
  );
  const elements = tables
    .flatMap((table) => table.elements ?? [])
    .filter((child): child is Element => child.type === "element" && child.name === "office:forms");
  const parsed = elements.map((element) => parseOfficeForms(element));
  if (!parsed.length) return undefined;
  return {
    ...parsed.reduce((left, right) => ({ ...left, ...right, containers: [] })),
    containers: parsed.flatMap((forms) => forms.containers),
    ...(parsed.some((forms) => forms.models)
      ? { models: parsed.flatMap((forms) => forms.models ?? []) }
      : {}),
  };
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
              "table:value-type": "date",
              "table:date-value": `${String(calculation.nullDate.year).padStart(4, "0")}-${String(
                calculation.nullDate.month,
              ).padStart(2, "0")}-${String(calculation.nullDate.day).padStart(2, "0")}`,
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

export function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/></office:document-styles>`;
}

export function rejectUnknownSpreadsheetChildren(
  body: Element | undefined,
  hasDataMappings = false,
): void {
  const allowed = new Set([
    "table:table",
    "table:calculation-settings",
    "office:forms",
    "table:named-expressions",
    "table:content-validations",
    "table:database-ranges",
    "table:data-pilot-tables",
    "table:consolidation",
    "table:label-ranges",
    "table:tracked-changes",
    ...(hasDataMappings ? ["calcext:data-mappings"] : []),
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
  const unknownAttributes = unknownAttributeNames(element, allowedAttributes);
  if (unknownAttributes.length) {
    throw unknownOdsElement(
      element,
      `/office:document-content/office:body${element.name === "form:form" ? "/office:spreadsheet/office:forms" : ""}`,
      unknownAttributes[0] ?? "",
      "attribute has no canonical mapping",
    );
  }
  const unknownChild = unknownChildElement(element, allowedChildren);
  if (unknownChild) {
    throw unknownOdsElement(
      element,
      `/office:document-content/office:body${element.name === "form:form" ? "/office:spreadsheet/office:forms" : ""}`,
      unknownChild.name ?? "",
      "child has no canonical mapping",
    );
  }
}
