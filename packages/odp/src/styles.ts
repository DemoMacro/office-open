import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

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
