import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  parseStyleMaps,
  parseStyleProperty,
  removeStylePropertyAttributes,
  styleMapsXml,
  stylePropertyXml,
  xmlElement,
} from "@office-open/odf";
import type { StyleFamily, StyleResourceOptions } from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";
import type { TextProperties } from "./text-run";

export function addTextStyle(properties: TextProperties, styles: string[]): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `C${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "style:text-underline-style": properties.underline ? "solid" : undefined,
        "fo:font-size": properties.size ? `${properties.size}pt` : undefined,
      }),
    ]),
  );
  return name;
}

export function styleResourcesXml(overlays: StyleResourceOptions[] | undefined): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement("style:style", { "style:name": overlay.name, "style:family": overlay.family }, [
      ...styleMapsXml(overlay.maps),
      ...overlay.properties.map((property) => stylePropertyXml(property, "emu")),
    ]),
  );
}

/** Column width style — pptx widths are EMU. */
export function addColumnStyle(width: number | string, styles: string[]): string {
  const name = `TC${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table-column" }, [
      xmlElement("style:table-column-properties", {
        "style:column-width": typeof width === "number" ? emuToLength(width) : width,
      }),
    ]),
  );
  return name;
}

export function parseTextStyles(container: Element | undefined): Map<string, TextProperties> {
  const result = new Map<string, TextProperties>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "text") continue;
    const properties = childNamed(style, "style:text-properties");
    const size = attributeString(properties, "fo:font-size");
    result.set(attributeString(style, "style:name") ?? "", {
      bold: attributeString(properties, "fo:font-weight") === "bold",
      italic: attributeString(properties, "fo:font-style") === "italic",
      underline:
        attributeString(properties, "style:text-underline-style") === "solid"
          ? "single"
          : undefined,
      size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
    });
  }
  return result;
}

const ODP_STYLE_PROPERTIES: Record<string, string[]> = {
  paragraph: ["style:paragraph-properties", "loext:graphic-properties", "style:text-properties"],
  text: ["style:text-properties"],
  graphic: [
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  presentation: [
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
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  section: ["style:section-properties"],
  "drawing-page": ["style:drawing-page-properties"],
  ruby: ["style:ruby-properties"],
};

export function parseStyleResources(container: Element | undefined): StyleResourceOptions[] {
  const result: StyleResourceOptions[] = [];
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const family = attributeString(style, "style:family") as StyleFamily;
    const path = `/office:document-content/office:automatic-styles/style:style[@style:name="${name}"]`;
    const allowed = Object.hasOwn(ODP_STYLE_PROPERTIES, family)
      ? ODP_STYLE_PROPERTIES[family]
      : undefined;
    if (!allowed) {
      const reason = "unknown style family";
      throw new OdpParseError(
        `content.xml: ${path}: ${family}: ${reason}`,
        "content.xml",
        path,
        family,
        reason,
      );
    }
    const properties: StyleResourceOptions["properties"] = [];
    const maps = parseStyleMaps(style);
    for (const property of style.elements ?? []) {
      if (property.type !== "element" || property.name === "style:map") continue;
      const parsed = parseStyleProperty(property, { family, scale: "emu" });
      if (!parsed.ok) {
        const name = parsed.attributeName ?? parsed.elementName;
        throw new OdpParseError(
          `content.xml: ${path}: ${name}: ${parsed.reason}`,
          "content.xml",
          path,
          name,
          parsed.reason,
        );
      }
      const value = removeStylePropertyAttributes(parsed.value, (attribute) =>
        isMappedTextAttribute(property.name ?? "", attribute),
      );
      if (hasStyleProjection(value)) properties.push(value);
    }
    if (properties.length > 0 || maps.length > 0) {
      result.push({ name, family, properties, ...(maps.length > 0 ? { maps } : {}) });
    }
  }
  return result;
}

function isMappedTextAttribute(element: string, attribute: string): boolean {
  return (
    element === "style:text-properties" &&
    ["fo:font-weight", "fo:font-style", "style:text-underline-style", "fo:font-size"].includes(
      attribute,
    )
  );
}

function hasStyleProjection(property: StyleResourceOptions["properties"][number]): boolean {
  return Object.keys(property.attributes).length > 0 || (property.children?.length ?? 0) > 0;
}
