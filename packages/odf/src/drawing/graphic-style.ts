/**
 * ODF graphic styles (style:family="graphic") shared by ODT and ODP shape
 * mapping: the draw:fill/draw:stroke tokens convert to the core fill and
 * outline options that docx wpsShape and pptx shape properties both accept.
 *
 * @module
 */

import type { FillOptions, OutlineOptions } from "@office-open/core/drawing";
import type { Element } from "@office-open/xml";

import { OdfSchemaError } from "../errors";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  xmlElement,
} from "../runtime/xml";

/** Graphic style attributes that map onto shape fill and outline. */
export interface GraphicStyle {
  fill?: string;
  fillColor?: string;
  stroke?: string;
  strokeWidth?: string;
  strokeColor?: string;
}

/** Parses the graphic-family styles from a style container. */
export function parseGraphicStyles(
  container: Element | undefined,
  family = "graphic",
): Map<string, GraphicStyle> {
  const result = new Map<string, GraphicStyle>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== family) continue;
    const name = attributeString(style, "style:name") ?? "";
    const properties = childNamed(style, "style:graphic-properties");
    result.set(name, {
      fill: attributeString(properties, "draw:fill"),
      fillColor: attributeString(properties, "draw:fill-color")?.replace("#", ""),
      stroke: attributeString(properties, "draw:stroke"),
      strokeWidth: attributeString(properties, "svg:stroke-width"),
      strokeColor: attributeString(properties, "svg:stroke-color")?.replace("#", ""),
    });
  }
  return result;
}

/** Solid color options reduce to the sRGB hex ODF attributes accept. */
export function hexColorValue(color: unknown): string | undefined {
  if (typeof color === "string") return color;
  if (typeof color === "object" && color !== null && "value" in color)
    return String((color as { value: unknown }).value);
  return undefined;
}

/** Adds the # prefix ODF color attributes carry. */
export function odfColor(hex: string | undefined): string | undefined {
  return hex ? `#${hex}` : undefined;
}

/** draw:fill/draw:fill-color → core fill options. */
export function graphicFill(graphic: GraphicStyle | undefined): FillOptions | undefined {
  if (graphic?.fill === "none") return { type: "none" };
  if (graphic?.fillColor) return { type: "solid", color: graphic.fillColor };
  return undefined;
}

/** draw:stroke/svg:stroke-* → core outline options. */
export function graphicOutline(graphic: GraphicStyle | undefined): OutlineOptions | undefined {
  if (!graphic) return undefined;
  if (graphic.stroke === "none") return { type: "noFill" };
  if (!graphic.strokeColor && !graphic.strokeWidth) return undefined;
  return {
    ...(graphic.strokeWidth ? { width: lengthToEmu(graphic.strokeWidth) ?? 0 } : {}),
    ...(graphic.strokeColor ? { type: "solidFill", color: graphic.strokeColor } : {}),
  };
}

/** Shape fill/outline land in a reusable graphic style pushed onto the styles buffer. */
export function pushShapeStyle(
  fill: FillOptions | null | undefined,
  outline: OutlineOptions | undefined,
  styles: string[],
  family = "graphic",
): string | undefined {
  if (family === "chart") rejectUnsupportedChartStyle(fill, outline);
  const fillColor =
    fill !== undefined && fill !== null && typeof fill === "object" && fill.type === "solid"
      ? hexColorValue(fill.color)
      : undefined;
  const attributes = {
    "draw:fill":
      fill !== null && typeof fill === "object" && fill.type === "none"
        ? "none"
        : fillColor
          ? "solid"
          : undefined,
    "draw:fill-color": odfColor(fillColor),
    "draw:stroke": outline?.type === "noFill" ? "none" : outline?.color ? "solid" : undefined,
    "svg:stroke-width":
      outline?.width !== undefined
        ? typeof outline.width === "number"
          ? emuToLength(outline.width)
          : outline.width
        : undefined,
    "svg:stroke-color":
      outline?.color !== undefined ? odfColor(hexColorValue(outline.color)) : undefined,
  };
  if (Object.values(attributes).every((value) => value === undefined)) return undefined;
  const name = `gr${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": family }, [
      xmlElement("style:graphic-properties", attributes),
    ]),
  );
  return name;
}

function rejectUnsupportedChartStyle(
  fill: FillOptions | null | undefined,
  outline: OutlineOptions | undefined,
): void {
  const unsupportedFill =
    typeof fill === "object" && fill !== null && !["solid", "none"].includes(fill.type);
  const unsupportedOutline =
    outline?.type !== undefined && !["solidFill", "noFill"].includes(outline.type);
  if (unsupportedFill || unsupportedOutline) {
    const style = unsupportedFill ? "fill" : "outline";
    const path =
      "/office:document-content/office:automatic-styles/style:style/style:graphic-properties";
    const reason = `${style} has no canonical ChartSpaceOptions mapping`;
    throw new OdfSchemaError(
      `content.xml: ${path}: style:graphic-properties: ${reason}`,
      "content.xml",
      path,
      "style:graphic-properties",
      reason,
    );
  }
}
