import type { SectionChild, SectionOptions } from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  parseStyleProperty,
  readXml,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";
import type { StyleColumnsOptions, StylePropertyOptions } from "@office-open/odf";
import type { XmlAttributes } from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdtGenerateError, OdtParseError } from "./error";
import type { OdtMasterPageOptions, OdtPageLayoutOptions } from "./options";

export interface MasterHeaderFooterContext {
  renderBlocks: (children: SectionChild[]) => string;
}

/** Section headers/footers render as master-page style:header/style:footer. */
export function masterHeaderFooter(
  section: Pick<SectionOptions, "headers" | "footers"> | undefined,
  context: MasterHeaderFooterContext,
): string {
  const header = section?.headers;
  const footer = section?.footers;
  if (
    (header?.default !== undefined && header.odd !== undefined) ||
    (footer?.default !== undefined && footer.odd !== undefined)
  ) {
    throw new OdtGenerateError(
      "ODT has no separate odd and default master-page header/footer regions",
      "styles.xml",
      "/office:document-styles/office:master-styles/style:master-page",
      "style:master-page",
      "ambiguous-master-header-footer-slots",
    );
  }
  return (
    headerFooterXml(header?.default ?? header?.odd, "style:header", context) +
    headerFooterXml(header?.even, "style:header-left", context) +
    headerFooterXml(header?.first, "style:header-first", context) +
    headerFooterXml(footer?.default ?? footer?.odd, "style:footer", context) +
    headerFooterXml(footer?.even, "style:footer-left", context) +
    headerFooterXml(footer?.first, "style:footer-first", context)
  );
}

export function headerFooterXml(
  children: SectionChild[] | undefined,
  element: string,
  context: MasterHeaderFooterContext,
): string {
  if (children === undefined) return "";
  return xmlElement(element, undefined, [context.renderBlocks(children)]);
}

/** Master-page header/footer regions → canonical section slots. */
export function parseMasterHeaderFooter(
  files: OdfFiles,
  parseBlocks: (elements: Element[], path: string) => SectionChild[],
): {
  headers?: SectionOptions["headers"];
  footers?: SectionOptions["footers"];
} {
  const master = childNamed(
    childNamed(readXml(files, "styles.xml"), "office:master-styles"),
    "style:master-page",
  );
  return master ? parseMasterRegions(master, parseBlocks) : {};
}

export function parseMasterPages(
  files: OdfFiles,
  parseBlocks: (elements: Element[], path: string) => SectionChild[],
): OdtMasterPageOptions[] {
  const styles = readXml(files, "styles.xml");
  return childrenNamed(childNamed(styles, "office:master-styles"), "style:master-page").flatMap(
    (master) => {
      const name = attributeString(master, "style:name");
      if (name === undefined) return [];
      const regions = parseMasterRegions(master, parseBlocks);
      return [
        {
          name,
          ...(attributeString(master, "style:page-layout-name")
            ? { pageLayoutName: attributeString(master, "style:page-layout-name") }
            : {}),
          ...regions,
        },
      ];
    },
  );
}

function parseMasterRegions(
  master: Element,
  parseBlocks: (elements: Element[], path: string) => SectionChild[],
): {
  headers?: SectionOptions["headers"];
  footers?: SectionOptions["footers"];
} {
  const regions = [
    ["style:header", "headers", "default"],
    ["style:header-left", "headers", "even"],
    ["style:header-first", "headers", "first"],
    ["style:footer", "footers", "default"],
    ["style:footer-left", "footers", "even"],
    ["style:footer-first", "footers", "first"],
  ] as const;
  const result: {
    headers?: SectionOptions["headers"];
    footers?: SectionOptions["footers"];
  } = {};
  for (const [elementName, family, slot] of regions) {
    const region = childNamed(master, elementName);
    if (!region) continue;
    const display = attributeString(region, "style:display");
    if (display !== undefined && display !== "true" && display !== "1") {
      throw new OdtParseError(
        `styles.xml: ${elementName} display state has no canonical section mapping`,
        "styles.xml",
        `/office:document-styles/office:master-styles/style:master-page/${elementName}/@style:display`,
        elementName,
        "unsupported-master-header-footer-display",
      );
    }
    const blocks = (region.elements ?? []).filter((child) => child.type === "element");
    const parsed = parseBlocks(
      blocks,
      `/office:document-styles/office:master-styles/style:master-page/${elementName}`,
    );
    (result[family] ??= {})[slot] = parsed;
  }
  return result;
}

export function parsePageLayouts(files: OdfFiles): OdtPageLayoutOptions[] {
  const styles = readXml(files, "styles.xml");
  return childrenNamed(childNamed(styles, "office:automatic-styles"), "style:page-layout").flatMap(
    (layout) => {
      const name = attributeString(layout, "style:name");
      if (name === undefined) return [];
      return [
        {
          name,
          ...parsePageLayoutDefinition(childNamed(layout, "style:page-layout-properties")),
        },
      ];
    },
  );
}

export function parsePageLayout(files: OdfFiles): SectionOptions["properties"] {
  const styles = readXml(files, "styles.xml");
  const master = childNamed(childNamed(styles, "office:master-styles"), "style:master-page");
  const pageLayoutName = attributeString(master, "style:page-layout-name");
  const pageLayouts = childrenNamed(
    childNamed(styles, "office:automatic-styles"),
    "style:page-layout",
  );
  const layout = childNamed(
    pageLayouts.find((item) => attributeString(item, "style:name") === pageLayoutName) ??
      pageLayouts[0],
    "style:page-layout-properties",
  );
  const definition = parsePageLayoutDefinition(layout);
  return definition.properties;
}

function parsePageLayoutDefinition(
  layout: Element | undefined,
): Pick<OdtPageLayoutOptions, "properties" | "columns" | "pageLayoutProperties"> {
  const width = lengthToTwips(attributeString(layout, "fo:page-width"));
  const height = lengthToTwips(attributeString(layout, "fo:page-height"));
  const orientation = attributeString(layout, "style:print-orientation");
  if (orientation !== undefined && orientation !== "portrait" && orientation !== "landscape") {
    throw new OdtParseError(
      `styles.xml: unsupported page orientation "${orientation}"`,
      "styles.xml",
      "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@style:print-orientation",
      "style:print-orientation",
      "unsupported-page-orientation",
    );
  }
  const writingMode = attributeString(layout, "style:writing-mode");
  const writingProperties = writingModeProperties(writingMode);
  const margin = {
    top: lengthToTwips(attributeString(layout, "fo:margin-top")),
    right: lengthToTwips(attributeString(layout, "fo:margin-right")),
    bottom: lengthToTwips(attributeString(layout, "fo:margin-bottom")),
    left: lengthToTwips(attributeString(layout, "fo:margin-left")),
  };
  const hasMargin = Object.values(margin).some((value) => value !== undefined);
  const columns = parseColumns(childNamed(layout, "style:columns"));
  const pageLayoutProperties = parsePageLayoutProperties(layout);
  const properties: SectionOptions["properties"] =
    width === undefined &&
    height === undefined &&
    !hasMargin &&
    !orientation &&
    columns === undefined
      ? undefined
      : {
          pageSize:
            width !== undefined || height !== undefined || orientation
              ? {
                  width,
                  height,
                  orientation: orientation as "portrait" | "landscape" | undefined,
                }
              : false,
          pageMargin: hasMargin ? margin : undefined,
          ...writingProperties,
          ...columnsSectionProperties(columns),
        };
  return {
    properties,
    ...(columns === undefined ? {} : { columns }),
    ...(pageLayoutProperties === undefined ? {} : { pageLayoutProperties }),
  };
}

function parsePageLayoutProperties(
  element: Element | undefined,
): OdtPageLayoutOptions["pageLayoutProperties"] {
  if (element === undefined) return undefined;
  const parsed = parseStyleProperty(element, {
    family: "page-layout",
    scale: "hundredth-millimeter",
  });
  if (!parsed.ok) {
    throw new OdtParseError(
      `styles.xml: ${parsed.reason}`,
      "styles.xml",
      "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties",
      parsed.attributeName ?? parsed.elementName,
      "invalid-page-layout-properties",
    );
  }
  return pageLayoutProperty(parsed.value);
}

function pageLayoutProperty(
  property: StylePropertyOptions,
): OdtPageLayoutOptions["pageLayoutProperties"] {
  if (property.name !== "style:page-layout-properties") {
    throw new OdtParseError(
      "styles.xml: page layout property has the wrong element name",
      "styles.xml",
      "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties",
      property.name,
      "invalid-page-layout-properties",
    );
  }
  return property;
}

function writingModeProperties(
  value: string | undefined,
): Pick<NonNullable<SectionOptions["properties"]>, "bidi" | "textDirection"> {
  if (value === undefined || value === "lr-tb" || value === "lr")
    return value === undefined ? {} : { textDirection: "lrTb" };
  if (value === "tb-rl") return { textDirection: "tbRl" };
  if (value === "rl-tb") return { bidi: true };
  throw new OdtParseError(
    `styles.xml: writing mode "${value}" has no canonical section mapping`,
    "styles.xml",
    "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@style:writing-mode",
    "style:writing-mode",
    "unsupported-writing-mode",
  );
}

function parseColumns(element: Element | undefined): StyleColumnsOptions | undefined {
  if (element === undefined) return undefined;
  const attributes: StyleColumnsOptions["attributes"] = {};
  const count = attributeString(element, "fo:column-count");
  if (count !== undefined) {
    const value = Number(count);
    if (!Number.isInteger(value) || value < 1) {
      throw new OdtParseError(
        `styles.xml: invalid column count "${count}"`,
        "styles.xml",
        "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/style:columns/@fo:column-count",
        "fo:column-count",
        "invalid-column-count",
      );
    }
    attributes["fo:column-count"] = value;
  }
  const gap = attributeString(element, "fo:column-gap");
  if (gap !== undefined) {
    const value = lengthToTwips(gap);
    if (value === undefined) throw invalidColumnAttribute("fo:column-gap", gap);
    attributes["fo:column-gap"] = value;
  }
  if (attributes["fo:column-count"] === undefined)
    throw invalidColumnAttribute("fo:column-count", "missing");
  const columns = (element.elements ?? [])
    .filter((child) => child.type === "element" && child.name === "style:column")
    .map((child) => ({
      attributes: Object.fromEntries(
        Object.entries(child.attributes ?? {}).map(([name, value]) => [name, String(value)]),
      ) as StyleColumnsOptions["columns"][number]["attributes"],
    }));
  const separator = childNamed(element, "style:column-sep");
  return {
    name: "style:columns",
    attributes,
    columns,
    ...(separator
      ? {
          separator: {
            attributes: Object.fromEntries(
              Object.entries(separator.attributes ?? {}).map(([name, value]) => [
                name,
                String(value),
              ]),
            ) as NonNullable<StyleColumnsOptions["separator"]>["attributes"],
          },
        }
      : {}),
  };
}

function columnsSectionProperties(
  columns: StyleColumnsOptions | undefined,
): Pick<NonNullable<SectionOptions["properties"]>, "columns"> {
  if (columns === undefined) return {};
  const count = columns.attributes["fo:column-count"];
  const gap = columns.attributes["fo:column-gap"];
  const separator = columns.separator?.attributes["style:style"];
  return {
    columns: {
      ...(typeof count === "number" ? { count } : {}),
      ...(typeof gap === "number" ? { space: gap } : {}),
      ...(separator === undefined ? {} : { separate: separator !== "none" }),
    },
  };
}

function invalidColumnAttribute(name: string, value: string): OdtParseError {
  return new OdtParseError(
    `styles.xml: invalid column attribute ${name}="${value}"`,
    "styles.xml",
    "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/style:columns",
    name,
    "invalid-column-attribute",
  );
}

export function pageLayoutColumnsXml(
  properties: SectionOptions["properties"],
  preserved?: StyleColumnsOptions,
): string {
  const columns = pageLayoutColumnsOptions(properties, preserved);
  if (columns === undefined) return "";
  return xmlElement("style:columns", columnAttributes(columns.attributes), [
    ...(columns.separator === undefined
      ? []
      : [xmlElement("style:column-sep", columnAttributes(columns.separator.attributes))]),
    ...columns.columns.map((column) =>
      xmlElement("style:column", columnAttributes(column.attributes)),
    ),
  ]);
}

export function pageLayoutColumnsOptions(
  properties: SectionOptions["properties"],
  preserved?: StyleColumnsOptions,
): StyleColumnsOptions | undefined {
  const canonical = properties?.columns;
  if (canonical === undefined && preserved === undefined) return undefined;
  const columns = preserved ?? { name: "style:columns" as const, attributes: {}, columns: [] };
  const attributes = { ...columns.attributes };
  if (canonical?.count !== undefined) attributes["fo:column-count"] = canonical.count;
  if (attributes["fo:column-count"] === undefined)
    attributes["fo:column-count"] = Math.max(columns.columns.length, 1);
  if (canonical?.space !== undefined) attributes["fo:column-gap"] = canonical.space;
  let separator = columns.separator;
  if (canonical?.separate !== undefined) {
    separator = canonical.separate
      ? (separator ?? { attributes: { "style:style": "solid", "style:width": 1 } })
      : undefined;
  }
  return { ...columns, attributes, ...(separator === undefined ? {} : { separator }) };
}

function columnAttributes(
  attributes: Partial<Record<string, string | number | boolean>>,
): XmlAttributes {
  return Object.fromEntries(
    Object.entries(attributes).map(([name, value]) => [
      name,
      typeof value !== "number"
        ? value
        : name === "style:height"
          ? `${value}%`
          : name === "fo:column-count"
            ? value
            : twipsToLength(value),
    ]),
  );
}

export function twipsToLength(value: number | string): string {
  return emuToLength(typeof value === "number" ? value * 635 : lengthToEmu(value));
}

export function lengthToTwips(value: string | undefined): number | undefined {
  const emu = lengthToEmu(value);
  return emu === undefined ? undefined : Math.round(emu / 635);
}
