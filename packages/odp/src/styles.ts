import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";
import type { OdpStyleOverlay } from "./semantics";
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

export function styleOverlaysXml(overlays: OdpStyleOverlay[] | undefined): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement(
      "style:style",
      { "style:name": overlay.name, "style:family": overlay.family },
      overlay.properties.map((property) =>
        xmlElement(property.name, overlayAttributesXml(property.attributes)),
      ),
    ),
  );
}

function overlayAttributesXml(
  attributes: Record<string, string | number | boolean>,
): Record<string, string | number | boolean> {
  return Object.fromEntries(
    Object.entries(attributes).map(([name, value]) => [
      name,
      typeof value === "string"
        ? value
        : typeof value === "boolean"
          ? value
          : isPercentAttribute(name)
            ? `${value}%`
            : isLengthAttribute(name)
              ? emuToLength(value)
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
  presentation: ["style:graphic-properties", "style:paragraph-properties", "style:text-properties"],
  table: ["style:table-properties"],
  "table-column": ["style:table-column-properties"],
  "table-row": ["style:table-row-properties"],
  "table-cell": [
    "style:table-cell-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  section: ["style:section-properties"],
  "drawing-page": ["style:drawing-page-properties"],
};

export function parseStyleOverlays(container: Element | undefined): OdpStyleOverlay[] {
  const result: OdpStyleOverlay[] = [];
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const family = attributeString(style, "style:family") ?? "";
    const path = `/office:document-content/office:automatic-styles/style:style[@style:name="${name}"]`;
    const allowed = ODP_STYLE_PROPERTIES[family];
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
    const properties: OdpStyleOverlay["properties"] = [];
    for (const property of style.elements ?? []) {
      if (property.type !== "element" || property.name === "style:tab-stops") continue;
      if (!allowed.includes(property.name ?? "")) {
        const reason = "unknown style property element";
        throw new OdpParseError(
          `content.xml: ${path}: ${property.name}: ${reason}`,
          "content.xml",
          path,
          property.name ?? "",
          reason,
        );
      }
      const attributes: Record<string, string | number> = {};
      for (const [attribute, value] of Object.entries(property.attributes ?? {})) {
        if (
          !/^(fo|style|text|draw|table|svg|dr3d|presentation|smil|officeooo|loext):/.test(attribute)
        ) {
          const reason = "attribute has no ODF RNG mapping";
          throw new OdpParseError(
            `content.xml: ${path}: ${attribute}: ${reason}`,
            "content.xml",
            path,
            attribute,
            reason,
          );
        }
        if (!isMappedTextAttribute(property.name ?? "", attribute))
          attributes[attribute] = overlayValue(attribute, String(value));
      }
      if (Object.keys(attributes).length > 0)
        properties.push({ name: property.name ?? "", attributes });
    }
    if (properties.length > 0) result.push({ name, family, properties });
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

function overlayValue(name: string, value: string): string | number {
  if (/^\d+(?:\.\d+)?%$/.test(value)) return Number.parseFloat(value);
  if (name === "style:column-width") return lengthToEmu(value) ?? value;
  if (/(angle|rotation)/.test(name) && Number.isFinite(Number(value))) return Number(value);
  return value;
}
