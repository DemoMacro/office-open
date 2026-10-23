import type { RunOptions } from "@office-open/docx";
import type { StyleFamily, StyleResourceOptions, StylePropertyOptions } from "@office-open/odf";
import {
  attributeString,
  childNamed,
  childrenNamed,
  lengthToEmu,
  parseStyleProperty,
  stylePropertyAttributes,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdtParseError } from "./error";
import { parseTabStop, tabStopXml, type TabStop } from "./paragraph";

export type OdtAutomaticStyleResource = StyleResourceOptions;

export type StyleMap = Map<
  string,
  {
    alignment?: string;
    columnWidth?: number;
    pageBreakBefore?: boolean;
    tabStops?: TabStop[];
    family: StyleFamily;
    properties?: StylePropertyOptions[];
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
    const resourceProperties = propertyElements.map((property) => {
      const parsed = parseStyleProperty(property, { family, scale: "hundredth-millimeter" });
      if (!parsed.ok)
        throw styleError(style, parsed.attributeName ?? parsed.elementName, parsed.reason);
      return parsed.value;
    });
    result.set(name, {
      alignment: attributeString(paragraph, "fo:text-align"),
      pageBreakBefore: attributeString(paragraph, "fo:break-before") === "page",
      tabStops: tabStops.length > 0 ? tabStops : undefined,
      columnWidth,
      family,
      ...(resourceProperties.some((property) => Object.keys(property.attributes).length > 0)
        ? { properties: resourceProperties }
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

export function automaticStyleResourcesXml(
  overlays: OdtAutomaticStyleResource[] | undefined,
): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement(
      "style:style",
      { "style:name": overlay.name, "style:family": overlay.family },
      overlay.properties
        .filter((property) => Object.keys(property.attributes).length > 0)
        .map((property) =>
          xmlElement(property.name, stylePropertyAttributes(property, "hundredth-millimeter")),
        ),
    ),
  );
}

function parseFamily(value: string | undefined, style: Element): StyleFamily {
  const families: StyleFamily[] = [
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
    "ruby",
  ];
  if (value && families.includes(value as StyleFamily)) return value as StyleFamily;
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
