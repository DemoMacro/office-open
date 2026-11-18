import type { RunOptions } from "@office-open/docx";
import type { StyleFamily, StyleResourceOptions, StylePropertyOptions } from "@office-open/odf";
import {
  attributeString,
  childNamed,
  childrenNamed,
  lengthToEmu,
  parseStyleProperty,
  stylePropertyXml,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdtParseError } from "./error";
import { parseTabStop, tabStopXml, type TabStop } from "./paragraph";

export type OdtAutomaticStyleResource = StyleResourceOptions;

export interface ParsedStyle {
  basedOn?: string;
  alignment?: string;
  columnWidth?: number;
  pageBreakBefore?: boolean;
  tabStops?: TabStop[];
  family: StyleFamily;
  properties?: StylePropertyOptions[];
  character: CharacterProperties;
}

export type StyleMap = Map<string, ParsedStyle>;

function canonicalAlignment(value: string | undefined): string | undefined {
  return value === "justify" ? "both" : value;
}

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
            "fo:text-align": properties.alignment === "both" ? "justify" : properties.alignment,
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

export function parseStyles(...containers: (Element | undefined)[]): StyleMap {
  const direct: StyleMap = new Map();
  const sourceElements = new Map<string, Element>();
  for (const container of containers) {
    for (const style of childrenNamed(container, "style:style")) {
      const name = attributeString(style, "style:name") ?? "";
      const family = parseFamily(attributeString(style, "style:family"), style);
      const propertyElements = (style.elements ?? []).filter(
        (child) => child.type === "element" && child.name !== "style:tab-stops",
      );
      const paragraph = childNamed(style, "style:paragraph-properties");
      const tabStops = childrenNamed(
        childNamed(paragraph, "style:tab-stops"),
        "style:tab-stop",
      ).map(parseTabStop);
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
      const breakBefore = attributeString(paragraph, "fo:break-before");
      const bold = attributeString(character, "fo:font-weight");
      const italic = attributeString(character, "fo:font-style");
      const underline = attributeString(character, "style:text-underline-style");
      const strike = attributeString(character, "style:text-line-through-style");
      const resourceProperties = propertyElements.map((property) => {
        const parsed = parseStyleProperty(property, { family, scale: "hundredth-millimeter" });
        if (!parsed.ok)
          throw styleError(style, parsed.attributeName ?? parsed.elementName, parsed.reason);
        return parsed.value;
      });
      const characterProperties: CharacterProperties = {
        ...(bold === undefined ? {} : { bold: bold === "bold" }),
        ...(italic === undefined ? {} : { italic: italic === "italic" }),
        ...(underline === undefined
          ? {}
          : { underline: underline === "solid" ? ("single" as const) : undefined }),
        ...(strike === undefined ? {} : { strike: strike === "solid" }),
        ...(size?.endsWith("pt") ? { size: Number(size.slice(0, -2)) } : {}),
        ...(color?.startsWith("#") ? { color: color.slice(1) } : {}),
        ...(attributeString(character, "fo:font-family")
          ? { font: attributeString(character, "fo:font-family") }
          : {}),
      };
      direct.set(name, {
        basedOn: attributeString(style, "style:parent-style-name"),
        alignment: canonicalAlignment(attributeString(paragraph, "fo:text-align")),
        pageBreakBefore: breakBefore === undefined ? undefined : breakBefore === "page",
        tabStops: tabStops.length > 0 ? tabStops : undefined,
        columnWidth,
        family,
        ...(resourceProperties.some(hasStyleProjection) ? { properties: resourceProperties } : {}),
        character: characterProperties,
      });
      sourceElements.set(name, style);
    }
  }
  const resolved: StyleMap = new Map();
  const resolve = (name: string, path: Set<string>): ParsedStyle | undefined => {
    const cached = resolved.get(name);
    if (cached) return cached;
    const style = direct.get(name);
    if (!style) return undefined;
    if (path.has(name))
      throw styleError(sourceElements.get(name)!, name, "cyclic parent style chain");
    const nextPath = new Set(path).add(name);
    const parent = style.basedOn ? resolve(style.basedOn, nextPath) : undefined;
    const merged = {
      ...parent,
      ...style,
      alignment: style.alignment ?? parent?.alignment,
      columnWidth: style.columnWidth ?? parent?.columnWidth,
      pageBreakBefore: style.pageBreakBefore ?? parent?.pageBreakBefore,
      tabStops: style.tabStops ?? parent?.tabStops,
      properties: [...(parent?.properties ?? []), ...(style.properties ?? [])],
      character: { ...parent?.character, ...style.character },
    };
    resolved.set(name, merged);
    return merged;
  };
  for (const name of direct.keys()) resolve(name, new Set());
  return resolved;
}

export function automaticStyleResourcesXml(
  overlays: OdtAutomaticStyleResource[] | undefined,
): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement(
      "style:style",
      { "style:name": overlay.name, "style:family": overlay.family },
      overlay.properties
        .filter(hasStyleProjection)
        .map((property) => stylePropertyXml(property, "hundredth-millimeter")),
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

function hasStyleProjection(property: StylePropertyOptions): boolean {
  return Object.keys(property.attributes).length > 0 || (property.children?.length ?? 0) > 0;
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
