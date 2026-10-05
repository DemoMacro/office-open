import type { RunOptions } from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdtParseError } from "./error";
import { parseTabStop, tabStopXml, type TabStop } from "./paragraph";

export type OdfStyleFamily =
  | "paragraph"
  | "text"
  | "graphic"
  | "table"
  | "table-column"
  | "table-row"
  | "table-cell"
  | "section"
  | "page-layout"
  | "drawing-page";

export interface OdtStylePropertyOverlay {
  /** RNG-valid property element retained after canonical fields are mapped. */
  name: string;
  attributes: Record<string, string | number>;
}

export interface OdtAutomaticStyleOverlay {
  name: string;
  family: OdfStyleFamily;
  properties: OdtStylePropertyOverlay[];
}

export type StyleMap = Map<
  string,
  {
    alignment?: string;
    columnWidth?: number;
    pageBreakBefore?: boolean;
    tabStops?: TabStop[];
    family: OdfStyleFamily;
    properties?: OdtStylePropertyOverlay[];
    character: CharacterProperties;
  }
>;

export interface CharacterProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: "single";
  strike?: boolean;
  size?: number;
  color?: string;
  font?: string;
}

export function characterProperties(run: RunOptions): CharacterProperties {
  return {
    bold: run.bold,
    italic: run.italic,
    underline: run.underline ? "single" : undefined,
    strike: run.strike,
    size: run.size,
    color:
      typeof run.color === "string" && /^[0-9A-Fa-f]{6}$/.test(run.color) ? run.color : undefined,
    font: typeof run.font === "string" ? run.font : undefined,
  };
}

export function addParagraphStyle(
  properties: { alignment?: string; pageBreakBefore?: boolean; tabStops?: TabStop[] },
  styles: string[],
): string {
  const name = `P${styles.length + 1}`;
  styles.push(
    xmlElement(
      "style:style",
      { "style:name": name, "style:family": "paragraph", "style:parent-style-name": "Standard" },
      [
        xmlElement(
          "style:paragraph-properties",
          {
            "fo:text-align": properties.alignment,
            "fo:break-before": properties.pageBreakBefore ? "page" : undefined,
          },
          properties.tabStops?.length
            ? [xmlElement("style:tab-stops", undefined, properties.tabStops.map(tabStopXml))]
            : undefined,
        ),
      ],
    ),
  );
  return name;
}

export function addCharacterStyle(
  properties: CharacterProperties,
  styles: string[],
): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `C${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "style:text-underline-style": properties.underline ? "solid" : undefined,
        "style:text-line-through-style": properties.strike ? "solid" : undefined,
        "fo:font-size": properties.size ? `${properties.size}pt` : undefined,
        "fo:color": properties.color ? `#${properties.color}` : undefined,
        "fo:font-family": properties.font,
      }),
    ]),
  );
  return name;
}

export function parseStyles(container: Element | undefined): StyleMap {
  const result: StyleMap = new Map();
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const family = parseFamily(attributeString(style, "style:family"), style);
    const propertyElements = (style.elements ?? []).filter(
      (child) => child.type === "element" && child.name !== "style:tab-stops",
    );
    const allowedProperties = FAMILY_PROPERTIES[family];
    for (const property of propertyElements) {
      if (!allowedProperties.includes(property.name ?? ""))
        throw styleError(style, property.name ?? "", "unknown style property element");
      for (const attribute of Object.keys(property.attributes ?? {})) {
        if (!ATTRIBUTE_PREFIXES[family].test(attribute))
          throw styleError(style, attribute, "attribute has no ODF RNG mapping");
      }
    }
    const paragraph = childNamed(style, "style:paragraph-properties");
    const tabStops = childrenNamed(childNamed(paragraph, "style:tab-stops"), "style:tab-stop").map(
      parseTabStop,
    );
    const character = childNamed(style, "style:text-properties");
    const columnWidthValue = attributeString(
      childNamed(style, "style:table-column-properties"),
      "style:column-width",
    );
    const columnWidth =
      columnWidthValue === undefined
        ? undefined
        : Math.round((lengthToEmu(columnWidthValue) ?? 0) / 635);
    const color = attributeString(character, "fo:color");
    const size = attributeString(character, "fo:font-size");
    const overlayProperties = propertyElements.map((property) => ({
      name: property.name ?? "",
      attributes: Object.fromEntries(
        Object.entries(property.attributes ?? {}).map(([attribute, value]) => [
          attribute,
          overlayValue(attribute, String(value)),
        ]),
      ),
    }));
    result.set(name, {
      alignment: attributeString(paragraph, "fo:text-align"),
      pageBreakBefore: attributeString(paragraph, "fo:break-before") === "page",
      tabStops: tabStops.length > 0 ? tabStops : undefined,
      columnWidth,
      family,
      ...(overlayProperties.some((property) => Object.keys(property.attributes).length > 0)
        ? { properties: overlayProperties }
        : {}),
      character: {
        bold: attributeString(character, "fo:font-weight") === "bold",
        italic: attributeString(character, "fo:font-style") === "italic",
        underline:
          attributeString(character, "style:text-underline-style") === "solid"
            ? "single"
            : undefined,
        strike: attributeString(character, "style:text-line-through-style") === "solid",
        size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
        color: color?.startsWith("#") ? color.slice(1) : undefined,
        font: attributeString(character, "fo:font-family"),
      },
    });
  }
  return result;
}

export function automaticStyleOverlaysXml(
  overlays: OdtAutomaticStyleOverlay[] | undefined,
): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement(
      "style:style",
      { "style:name": overlay.name, "style:family": overlay.family },
      overlay.properties
        .filter((property) => Object.keys(property.attributes).length > 0)
        .map((property) => xmlElement(property.name, overlayAttributesXml(property.attributes))),
    ),
  );
}

function overlayAttributesXml(
  attributes: Record<string, string | number>,
): Record<string, string | number> {
  return Object.fromEntries(
    Object.entries(attributes).map(([name, value]) => [
      name,
      typeof value === "string"
        ? value
        : isPercentAttribute(name)
          ? `${value}%`
          : isLengthAttribute(name)
            ? emuToLength(value * 635)
            : value,
    ]),
  );
}

function isLengthAttribute(name: string): boolean {
  return /(width|height|margin|spacing|distance|indent|padding|position)/.test(name);
}

function isPercentAttribute(name: string): boolean {
  return /(line-height|rel-width|rel-height|opacity|transparency)/.test(name);
}

const FAMILY_PROPERTIES: Record<OdfStyleFamily, string[]> = {
  paragraph: ["style:paragraph-properties", "loext:graphic-properties", "style:text-properties"],
  text: ["style:text-properties"],
  graphic: [
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  table: ["style:table-properties"],
  "table-column": ["style:table-column-properties"],
  "table-row": ["style:table-row-properties"],
  "table-cell": [
    "style:table-cell-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  section: ["style:section-properties"],
  "page-layout": ["style:page-layout-properties"],
  "drawing-page": ["style:drawing-page-properties"],
};

const ATTRIBUTE_PREFIXES: Record<OdfStyleFamily, RegExp> = Object.fromEntries([
  ["paragraph", /^(fo|style|text|svg|draw|officeooo|loext):/],
  ["text", /^(fo|style|text|officeooo|loext):/],
  ["graphic", /^(fo|style|draw|svg|dr3d|smil|officeooo|loext):/],
  ["table", /^(fo|style|table|officeooo|loext):/],
  ["table-column", /^(fo|style|officeooo|loext):/],
  ["table-row", /^(fo|style|officeooo|loext):/],
  ["table-cell", /^(fo|style|officeooo|loext):/],
  ["section", /^(fo|style|text|officeooo|loext):/],
  ["page-layout", /^(fo|style|text|svg|officeooo|loext):/],
  ["drawing-page", /^(fo|style|draw|svg|dr3d|presentation|smil|officeooo|loext):/],
]) as Record<OdfStyleFamily, RegExp>;

function parseFamily(value: string | undefined, style: Element): OdfStyleFamily {
  const families: OdfStyleFamily[] = [
    "paragraph",
    "text",
    "graphic",
    "table",
    "table-column",
    "table-row",
    "table-cell",
    "section",
    "page-layout",
    "drawing-page",
  ];
  if (value && families.includes(value as OdfStyleFamily)) return value as OdfStyleFamily;
  throw styleError(style, value ?? "style:family", "unknown style family");
}

function styleError(style: Element, name: string, reason: string): OdtParseError {
  const styleName = attributeString(style, "style:name") ?? "";
  const path = `/office:document-content/office:automatic-styles/style:style[@style:name="${styleName}"]`;
  return new OdtParseError(
    `content.xml: ${path}: ${name}: ${reason}`,
    "content.xml",
    path,
    name,
    reason,
  );
}

function overlayValue(name: string, value: string): string | number {
  if (/^\d+(?:\.\d+)?%$/.test(value)) return Number.parseFloat(value);
  const length = lengthToEmu(value);
  if (
    length !== undefined &&
    /(width|height|margin|spacing|distance|indent|padding|position)/.test(name)
  )
    return Math.round(length / 635);
  if (/(angle|rotation)/.test(name)) {
    const angle = Number(value);
    if (Number.isFinite(angle)) return angle;
  }
  return value;
}
