import type { CorePropertiesOptions } from "@office-open/core";
import type {
  AutoReloadOptions,
  CalendarDurationOptions,
  DocumentStatisticsOptions,
  DocumentTemplateOptions,
  HyperlinkBehaviorOptions,
  UserDefinedMetadataOptions,
} from "@office-open/core/opc";
import { escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { readXml, type OdfFiles } from "./container/package";
import { OdfSchemaError } from "./errors";
import { childNamed, textOf } from "./runtime/xml";

export const ODF_NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(" ");

export function metaXml(options: CorePropertiesOptions): string {
  const text = (value: string | null | undefined): boolean =>
    value !== undefined && value !== null && value !== "";
  const fields = [
    text(options.title) && `<dc:title>${escapeXml(options.title!)}</dc:title>`,
    text(options.subject) && `<dc:subject>${escapeXml(options.subject!)}</dc:subject>`,
    text(options.creator) &&
      `<meta:initial-creator>${escapeXml(options.creator!)}</meta:initial-creator>`,
    text(options.keywords) && `<meta:keyword>${escapeXml(options.keywords!)}</meta:keyword>`,
    text(options.description) &&
      `<dc:description>${escapeXml(options.description!)}</dc:description>`,
    text(options.lastModifiedBy) &&
      `<dc:creator>${escapeXml(options.lastModifiedBy!)}</dc:creator>`,
    text(options.created) &&
      `<meta:creation-date>${escapeXml(options.created!)}</meta:creation-date>`,
    text(options.modified) && `<dc:date>${escapeXml(options.modified!)}</dc:date>`,
    text(options.category) && `<office:category>${escapeXml(options.category!)}</office:category>`,
    text(options.language) && `<dc:language>${escapeXml(options.language!)}</dc:language>`,
    ...(options.publishers ?? []).map(
      (publisher) => `<dc:publisher>${escapeXml(publisher)}</dc:publisher>`,
    ),
    ...(options.autoFillers ?? []).map(
      (autoFiller) => `<dc:auto-filler>${escapeXml(autoFiller)}</dc:auto-filler>`,
    ),
    ...(options.sources ?? []).map((source) => `<dc:source>${escapeXml(source)}</dc:source>`),
    text(options.identifier) && `<dc:identifier>${escapeXml(options.identifier!)}</dc:identifier>`,
    text(options.lastPrinted) &&
      `<meta:print-date>${escapeXml(options.lastPrinted!)}</meta:print-date>`,
    text(options.printedBy) &&
      `<meta:printed-by>${escapeXml(options.printedBy!)}</meta:printed-by>`,
    text(options.application) &&
      `<meta:generator>${escapeXml(options.application!)}</meta:generator>`,
    options.revision !== undefined &&
      `<meta:editing-cycles>${options.revision}</meta:editing-cycles>`,
    options.editingDuration &&
      `<meta:editing-duration>${formatDuration(options.editingDuration)}</meta:editing-duration>`,
    templateXml(options.template),
    autoReloadXml(options.autoReload),
    hyperlinkBehaviorXml(options.hyperlinkBehavior),
    documentStatisticsXml(options.statistics),
    ...(options.userDefinedProperties ?? []).map(userDefinedXml),
  ].filter(Boolean);
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${ODF_NAMESPACES} office:version="1.3"><office:meta>${fields.join(
    "",
  )}</office:meta></office:document-meta>`;
}

const META_FIELDS: ReadonlySet<string> = new Set([
  "dc:title",
  "dc:subject",
  "dc:initial-creator",
  "meta:initial-creator",
  "dc:creator",
  "meta:keyword",
  "dc:description",
  "meta:creation-date",
  "dc:date",
  "office:category",
  "dc:language",
  "dc:publisher",
  "dc:auto-filler",
  "dc:source",
  "dc:identifier",
  "meta:print-date",
  "meta:printed-by",
  "meta:editing-cycles",
  "meta:editing-duration",
  "meta:generator",
  "meta:template",
  "meta:auto-reload",
  "meta:hyperlink-behaviour",
  "meta:document-statistic",
]);

export function parseMeta(files: OdfFiles): CorePropertiesOptions {
  if (!files["meta.xml"]) return {};
  const meta = childNamed(readXml(files, "meta.xml"), "office:meta");
  rejectUnknownMetadata(meta);
  const value = (name: string): string | undefined => textOf(childNamed(meta, name)) || undefined;
  const revision = value("meta:editing-cycles");
  const keywords = childrenNamed(meta, "meta:keyword").map(textOf).filter(Boolean).join(", ");
  const properties: CorePropertiesOptions = {
    title: value("dc:title"),
    subject: value("dc:subject"),
    creator: value("meta:initial-creator") ?? value("dc:initial-creator"),
    keywords: keywords || undefined,
    description: value("dc:description"),
    lastModifiedBy: value("dc:creator"),
    created: value("meta:creation-date"),
    modified: value("dc:date"),
    category: value("office:category"),
    language: value("dc:language"),
    publishers: childrenNamed(meta, "dc:publisher").map(textOf).filter(Boolean),
    autoFillers: childrenNamed(meta, "dc:auto-filler").map(textOf).filter(Boolean),
    sources: childrenNamed(meta, "dc:source").map(textOf).filter(Boolean),
    identifier: value("dc:identifier"),
    lastPrinted: value("meta:print-date"),
    printedBy: value("meta:printed-by"),
    application: value("meta:generator"),
    revision: revision ? Number(revision) : undefined,
    editingDuration: parseDuration(childNamed(meta, "meta:editing-duration")),
    template: parseTemplate(childNamed(meta, "meta:template")),
    autoReload: parseAutoReload(childNamed(meta, "meta:auto-reload")),
    hyperlinkBehavior: parseHyperlinkBehavior(childNamed(meta, "meta:hyperlink-behaviour")),
    statistics: parseDocumentStatistics(childNamed(meta, "meta:document-statistic")),
    userDefinedProperties: childrenNamed(meta, "meta:user-defined").map(parseUserDefined),
  };
  for (const [key, value] of Object.entries(properties)) {
    if (value === undefined || (Array.isArray(value) && value.length === 0)) {
      delete properties[key as keyof CorePropertiesOptions];
    }
  }
  return properties;
}

function childrenNamed(element: Element | undefined, name: string): Element[] {
  return element?.elements?.filter((child) => child.name === name) ?? [];
}

function metadataError(name: string, reason: string): OdfSchemaError {
  const path = `/office:document-meta/office:meta/${name}`;
  return new OdfSchemaError(`meta.xml: ${path}: ${reason}`, "meta.xml", path, name, reason);
}

function rejectUnknownMetadata(meta: Element | undefined): void {
  const unknownAttributes = Object.keys(meta?.attributes ?? {}).filter(
    (name) => !name.startsWith("xmlns:"),
  );
  if (unknownAttributes.length > 0) {
    const name = unknownAttributes[0] ?? "";
    throw metadataError(name, "attribute has no canonical mapping");
  }
  for (const child of meta?.elements ?? []) {
    if (child.type !== "element") continue;
    const repeated =
      child.name === "meta:user-defined" ||
      child.name === "meta:keyword" ||
      child.name === "dc:publisher" ||
      child.name === "dc:auto-filler" ||
      child.name === "dc:source";
    if (!child.name || (!repeated && !META_FIELDS.has(child.name))) {
      const name = child.name ?? "unknown";
      throw metadataError(name, "no canonical metadata mapping");
    }
    if (!repeated && childNamed(meta, child.name) !== child) {
      throw metadataError(child.name!, "repeated field has no canonical mapping");
    }
  }
}

function parseDuration(element: Element | undefined): CalendarDurationOptions | undefined {
  const value = textOf(element);
  return parseDurationValue(value);
}

function parseDurationValue(value: string): CalendarDurationOptions | undefined {
  if (!value) return undefined;
  const match =
    /^(-)?P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T(?=\d)(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(
      value,
    );
  const parts = match?.slice(2, 8).map((part) => (part === undefined ? 0 : Number(part))) ?? [];
  const hasComponent = match?.slice(2, 8).some((part) => part !== undefined) ?? false;
  if (!match || !hasComponent) {
    throw metadataError("duration", `invalid duration "${value}"`);
  }
  const duration: CalendarDurationOptions = {};
  if (match[1]) duration.negative = true;
  if (parts[0]) duration.years = parts[0];
  if (parts[1]) duration.months = parts[1];
  if (parts[2]) duration.days = parts[2];
  if (parts[3]) duration.hours = parts[3];
  if (parts[4]) duration.minutes = parts[4];
  if (parts[5]) duration.seconds = parts[5];
  return duration;
}

function formatDuration(duration: CalendarDurationOptions): string {
  const parts = [
    duration.years,
    duration.months,
    duration.days,
    duration.hours,
    duration.minutes,
    duration.seconds,
  ];
  if (parts.some((part) => part !== undefined && (!Number.isFinite(part) || part < 0))) {
    throw metadataError("duration", "duration components must be non-negative");
  }
  const date = [
    duration.years && `${duration.years}Y`,
    duration.months && `${duration.months}M`,
    duration.days && `${duration.days}D`,
  ]
    .filter(Boolean)
    .join("");
  const time = [
    duration.hours && `${duration.hours}H`,
    duration.minutes && `${duration.minutes}M`,
    duration.seconds !== undefined && `${duration.seconds}S`,
  ]
    .filter(Boolean)
    .join("");
  if (!date && !time) return "PT0S";
  const prefix = `${duration.negative ? "-" : ""}P${date}`;
  return time ? `${prefix}T${time}` : prefix;
}

function parseTemplate(element: Element | undefined): DocumentTemplateOptions | undefined {
  if (!element) return undefined;
  const href = attributeString(element, "xlink:href");
  if (href === undefined) throw metadataError("meta:template", "missing template href");
  return {
    href,
    actuate: attributeString(element, "xlink:actuate") as DocumentTemplateOptions["actuate"],
    title: attributeString(element, "xlink:title"),
    modified: attributeString(element, "meta:date"),
  };
}

function templateXml(template: DocumentTemplateOptions | undefined): string | false {
  if (!template) return false;
  return `<meta:template xlink:type="simple" xlink:href="${escapeXml(template.href)}"${
    template.actuate ? ` xlink:actuate="${template.actuate}"` : ""
  }${template.title ? ` xlink:title="${escapeXml(template.title)}"` : ""}${
    template.modified ? ` meta:date="${escapeXml(template.modified)}"` : ""
  }/>`;
}

function parseAutoReload(element: Element | undefined): AutoReloadOptions | undefined {
  if (!element) return undefined;
  return {
    href: attributeString(element, "xlink:href"),
    show: attributeString(element, "xlink:show") as AutoReloadOptions["show"],
    delay: parseDurationValue(attributeString(element, "meta:delay") ?? ""),
  };
}

function autoReloadXml(options: AutoReloadOptions | undefined): string | false {
  if (!options) return false;
  return `<meta:auto-reload${
    options.href ? ` xlink:type="simple" xlink:href="${escapeXml(options.href)}"` : ""
  }${options.show ? ` xlink:show="${options.show}"` : ""}${
    options.delay ? ` meta:delay="${formatDuration(options.delay)}"` : ""
  }/>`;
}

function parseHyperlinkBehavior(
  element: Element | undefined,
): HyperlinkBehaviorOptions | undefined {
  if (!element) return undefined;
  return {
    targetFrameName: attributeString(element, "office:target-frame-name"),
    show: attributeString(element, "xlink:show") as HyperlinkBehaviorOptions["show"],
  };
}

function hyperlinkBehaviorXml(options: HyperlinkBehaviorOptions | undefined): string | false {
  if (!options) return false;
  return `<meta:hyperlink-behaviour${
    options.targetFrameName
      ? ` office:target-frame-name="${escapeXml(options.targetFrameName)}"`
      : ""
  }${options.show ? ` xlink:show="${options.show}"` : ""}/>`;
}

const STATISTIC_FIELDS = [
  ["pageCount", "meta:page-count"],
  ["tableCount", "meta:table-count"],
  ["drawCount", "meta:draw-count"],
  ["imageCount", "meta:image-count"],
  ["oleObjectCount", "meta:ole-object-count"],
  ["objectCount", "meta:object-count"],
  ["paragraphCount", "meta:paragraph-count"],
  ["wordCount", "meta:word-count"],
  ["characterCount", "meta:character-count"],
  ["frameCount", "meta:frame-count"],
  ["sentenceCount", "meta:sentence-count"],
  ["syllableCount", "meta:syllable-count"],
  ["nonWhitespaceCharacterCount", "meta:non-whitespace-character-count"],
  ["rowCount", "meta:row-count"],
  ["cellCount", "meta:cell-count"],
] as const;

function statisticNumber(element: Element, name: string): number | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 0) {
    throw metadataError("meta:document-statistic", `invalid count for ${name}`);
  }
  return parsed;
}

function parseDocumentStatistics(
  element: Element | undefined,
): DocumentStatisticsOptions | undefined {
  if (!element) return undefined;
  const statistics: DocumentStatisticsOptions = {};
  for (const [key, name] of STATISTIC_FIELDS) {
    const value = statisticNumber(element, name);
    if (value !== undefined) statistics[key] = value;
  }
  return statistics;
}

function documentStatisticsXml(statistics: DocumentStatisticsOptions | undefined): string | false {
  if (!statistics) return false;
  const attributes = STATISTIC_FIELDS.filter(([key]) => statistics[key] !== undefined)
    .map(([key, name]) => ` ${name}="${statistics[key]}"`)
    .join("");
  if (!attributes) return false;
  return `<meta:document-statistic${attributes}/>`;
}

function parseUserDefined(element: Element): UserDefinedMetadataOptions {
  const name = attributeString(element, "meta:name");
  if (name === undefined) throw metadataError("meta:user-defined", "missing property name");
  const type = attributeString(element, "meta:value-type");
  const text = textOf(element);
  if (type === "float") {
    const value = Number(text);
    if (!Number.isFinite(value)) throw metadataError("meta:user-defined", "invalid float value");
    return { name, type: "float", value };
  }
  if (type === "date") return { name, type: "date", value: text };
  if (type === "time") {
    const value = parseDuration(element);
    if (!value) throw metadataError("meta:user-defined", "missing time value");
    return { name, type: "time", value };
  }
  if (type === "boolean") {
    if (text !== "true" && text !== "false") {
      throw metadataError("meta:user-defined", "invalid boolean value");
    }
    return { name, type: "boolean", value: text === "true" };
  }
  if (type === "string") return { name, type: "string", value: text };
  if (type === undefined) return { name, type: "text", value: text };
  throw metadataError("meta:user-defined", `unknown value type "${type}"`);
}

function userDefinedXml(property: UserDefinedMetadataOptions): string {
  const value = property.type === "time" ? formatDuration(property.value) : String(property.value);
  const typeAttribute = property.type === "text" ? "" : ` meta:value-type="${property.type}"`;
  return `<meta:user-defined meta:name="${escapeXml(property.name)}"${typeAttribute}>${escapeXml(value)}</meta:user-defined>`;
}

function attributeString(element: Element | undefined, name: string): string | undefined {
  const value = element?.attributes?.[name];
  return value === undefined ? undefined : String(value);
}
