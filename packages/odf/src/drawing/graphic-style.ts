/**
 * ODF graphic styles (style:family="graphic") shared by ODT and ODP shape
 * mapping: the draw:fill/draw:stroke tokens convert to the core fill and
 * outline options that docx wpsShape and pptx shape properties both accept.
 *
 * @module
 */

import { toUint8Array, type DataType } from "@office-open/core";
import type { FillOptions, OutlineOptions, SolidFillOptions } from "@office-open/core/drawing";
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
  fillGradientName?: string;
  fillHatchName?: string;
  fillImageName?: string;
  fillRepeat?: "no-repeat" | "repeat" | "stretch";
  fillOptions?: FillOptions;
  outlineOptions?: OutlineOptions;
}

export interface GraphicStyleResource {
  name: string;
  displayName?: string;
  attributes: Record<string, string>;
}

export interface GraphicFillImageResource extends GraphicStyleResource {
  data?: Uint8Array;
}

export interface GraphicStyleResources {
  gradients: GraphicStyleResource[];
  hatches: GraphicStyleResource[];
  fillImages: GraphicFillImageResource[];
}

export function createGraphicStyleResources(): GraphicStyleResources {
  return { gradients: [], hatches: [], fillImages: [] };
}

const styleResourceBuffers = new WeakMap<string[], GraphicStyleResources>();

export function bindGraphicStyleResources(
  styles: string[],
  resources: GraphicStyleResources,
): void {
  styleResourceBuffers.set(styles, resources);
}

export function graphicStyleResources(styles: string[]): GraphicStyleResources {
  const resources = styleResourceBuffers.get(styles);
  if (resources !== undefined) return resources;
  const created = createGraphicStyleResources();
  styleResourceBuffers.set(styles, created);
  return created;
}

export function parseGraphicStyleResources(
  containers: readonly (Element | undefined)[],
  binaries: Record<string, Uint8Array> = {},
): GraphicStyleResources {
  const resources = createGraphicStyleResources();
  for (const container of containers) {
    collectGraphicResource(container, "draw:gradient", "draw:name", resources.gradients);
    collectGraphicResource(container, "draw:hatch", "draw:name", resources.hatches);
    for (const element of childrenNamed(container, "draw:fill-image")) {
      const name = attributeString(element, "draw:name");
      if (!name || resources.fillImages.some((resource) => resource.name === name)) continue;
      const href = attributeString(element, "xlink:href");
      const path = href && !isExternalReference(href) ? normalizePackagePath(href) : undefined;
      resources.fillImages.push({
        name,
        displayName: attributeString(element, "draw:display-name"),
        attributes: withoutResourceAttributes(element, ["draw:name", "draw:display-name"]),
        ...(path && binaries[path] ? { data: binaries[path] } : {}),
      });
    }
  }
  return resources;
}

/** Parses the graphic-family styles from a style container. */
export function parseGraphicStyles(
  container: Element | undefined,
  family = "graphic",
  resources?: GraphicStyleResources,
  inherited?: ReadonlyMap<string, GraphicStyle>,
): Map<string, GraphicStyle> {
  const direct = new Map<string, { style: GraphicStyle; parent?: string }>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== family) continue;
    const name = attributeString(style, "style:name") ?? "";
    const properties = childNamed(style, "style:graphic-properties");
    const graphic: GraphicStyle = {
      fill: attributeString(properties, "draw:fill"),
      fillColor: attributeString(properties, "draw:fill-color")?.replace("#", ""),
      stroke: attributeString(properties, "draw:stroke"),
      strokeWidth: attributeString(properties, "svg:stroke-width"),
      strokeColor: attributeString(properties, "svg:stroke-color")?.replace("#", ""),
      fillGradientName: attributeString(properties, "draw:fill-gradient-name"),
      fillHatchName: attributeString(properties, "draw:fill-hatch-name"),
      fillImageName: attributeString(properties, "draw:fill-image-name"),
      fillRepeat: parseFillRepeat(attributeString(properties, "style:repeat")),
    };
    direct.set(name, {
      style: graphic,
      parent: attributeString(style, "style:parent-style-name"),
    });
  }
  const result = new Map<string, GraphicStyle>();
  const resolving = new Set<string>();
  const resolve = (name: string): GraphicStyle | undefined => {
    const cached = result.get(name);
    if (cached) return cached;
    const entry = direct.get(name);
    if (!entry) return inherited?.get(name);
    if (resolving.has(name))
      throw new OdfSchemaError(
        `Graphic style ${name} has a cyclic parent-style-name chain`,
        "styles.xml",
        `/office:document-styles/office:styles/style:style[@style:name="${name}"]`,
        "style:parent-style-name",
        "cyclic-style-inheritance",
      );
    resolving.add(name);
    const parent = entry.parent ? resolve(entry.parent) : undefined;
    const merged = mergeGraphicStyle(parent, entry.style);
    if (resources !== undefined) {
      const fill = projectGraphicFill(merged, resources);
      const outline = projectGraphicOutline(merged);
      if (fill !== undefined) merged.fillOptions = fill;
      if (outline !== undefined) merged.outlineOptions = outline;
    }
    resolving.delete(name);
    result.set(name, merged);
    return merged;
  };
  for (const name of direct.keys()) resolve(name);
  return result;
}

function mergeGraphicStyle(parent: GraphicStyle | undefined, child: GraphicStyle): GraphicStyle {
  return {
    fill: child.fill ?? parent?.fill,
    fillColor: child.fillColor ?? parent?.fillColor,
    stroke: child.stroke ?? parent?.stroke,
    strokeWidth: child.strokeWidth ?? parent?.strokeWidth,
    strokeColor: child.strokeColor ?? parent?.strokeColor,
    fillGradientName: child.fillGradientName ?? parent?.fillGradientName,
    fillHatchName: child.fillHatchName ?? parent?.fillHatchName,
    fillImageName: child.fillImageName ?? parent?.fillImageName,
    fillRepeat: child.fillRepeat ?? parent?.fillRepeat,
    fillOptions: child.fillOptions ?? parent?.fillOptions,
    outlineOptions: child.outlineOptions ?? parent?.outlineOptions,
  };
}

/** Solid color options reduce to the sRGB hex ODF attributes accept. */
export function hexColorValue(
  color: FillOptions | SolidFillOptions | string | null | undefined,
): string | undefined {
  if (typeof color === "string") return color;
  if (color !== null && color !== undefined && "value" in color && typeof color.value === "string")
    return color.value;
  return undefined;
}

/** Adds the # prefix ODF color attributes carry. */
export function odfColor(hex: string | undefined): string | undefined {
  return hex ? `#${hex}` : undefined;
}

/** draw:fill/draw:fill-color → core fill options. */
export function graphicFill(graphic: GraphicStyle | undefined): FillOptions | undefined {
  if (graphic?.fillOptions !== undefined) return graphic.fillOptions;
  if (graphic?.fill === "none") return { type: "none" };
  if (graphic?.fillColor) return { type: "solid", color: graphic.fillColor };
  return undefined;
}

/** draw:stroke/svg:stroke-* → core outline options. */
export function graphicOutline(graphic: GraphicStyle | undefined): OutlineOptions | undefined {
  if (graphic?.outlineOptions !== undefined) return graphic.outlineOptions;
  if (!graphic) return undefined;
  if (graphic.stroke === "none") return { type: "noFill" };
  if (!graphic.strokeColor && !graphic.strokeWidth) return undefined;
  return {
    ...(graphic.strokeWidth ? { width: lengthToEmu(graphic.strokeWidth) ?? 0 } : {}),
    ...(graphic.strokeColor ? { type: "solidFill", color: graphic.strokeColor } : {}),
  };
}

function parseFillRepeat(value: string | undefined): GraphicStyle["fillRepeat"] {
  if (value === "no-repeat" || value === "repeat" || value === "stretch") return value;
  return undefined;
}

/** Shape fill/outline land in a reusable graphic style pushed onto the styles buffer. */
export function pushShapeStyle(
  fill: FillOptions | null | undefined,
  outline: OutlineOptions | undefined,
  styles: string[],
  family = "graphic",
): string | undefined {
  const resources = graphicStyleResources(styles);
  const fillResource = fillResourceXml(fill, resources);
  if (family === "chart") rejectUnsupportedChartStyle(fill, outline);
  const fillColor =
    fill !== undefined && fill !== null && typeof fill === "object" && fill.type === "solid"
      ? hexColorValue(fill.color)
      : undefined;
  const attributes = {
    "draw:fill":
      fill !== null && typeof fill === "object" && fill.type === "none"
        ? "none"
        : fillResource?.type
          ? fillResource.type
          : fillColor
            ? "solid"
            : undefined,
    "draw:fill-color": odfColor(fillColor),
    ...fillResource?.attributes,
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

export function graphicStyleResourcesXml(resources: GraphicStyleResources): string[] {
  return [
    ...resources.gradients.map((resource) => resourceXml("draw:gradient", resource)),
    ...resources.hatches.map((resource) => resourceXml("draw:hatch", resource)),
    ...resources.fillImages.map((resource) => resourceXml("draw:fill-image", resource)),
  ];
}

export function graphicStyleFiles(resources: GraphicStyleResources): {
  path: string;
  data: Uint8Array;
}[] {
  return resources.fillImages.flatMap((resource) => {
    const href = resource.attributes["xlink:href"];
    if (!href || isExternalReference(href) || resource.data === undefined) return [];
    return [{ path: normalizePackagePath(href), data: toUint8Array(resource.data) }];
  });
}

function collectGraphicResource(
  container: Element | undefined,
  name: string,
  nameAttribute: string,
  output: GraphicStyleResource[],
): void {
  for (const element of childrenNamed(container, name)) {
    const resourceName = attributeString(element, nameAttribute);
    if (!resourceName || output.some((resource) => resource.name === resourceName)) continue;
    output.push({
      name: resourceName,
      displayName: attributeString(element, "draw:display-name"),
      attributes: withoutResourceAttributes(element, [nameAttribute, "draw:display-name"]),
    });
  }
}

function withoutResourceAttributes(
  element: Element,
  excluded: readonly string[],
): Record<string, string> {
  return Object.fromEntries(
    Object.entries(element.attributes ?? {})
      .filter(([name]) => !excluded.includes(name))
      .map(([name, value]) => [name, String(value)]),
  );
}

function resourceXml(
  elementName: "draw:gradient" | "draw:hatch" | "draw:fill-image",
  resource: GraphicStyleResource,
): string {
  return xmlElement(elementName, {
    "draw:name": resource.name,
    "draw:display-name": resource.displayName,
    ...resource.attributes,
    ...(elementName === "draw:fill-image"
      ? { "xlink:type": "simple", "xlink:show": "embed", "xlink:actuate": "onLoad" }
      : {}),
  });
}

function projectGraphicFill(
  graphic: GraphicStyle,
  resources: GraphicStyleResources,
): FillOptions | undefined {
  const token = graphic.fill;
  if (token === undefined) return undefined;
  if (token === "none") return { type: "none" };
  if (token === "solid") return { type: "solid", color: graphic.fillColor ?? "000000" };
  if (token === "gradient") {
    const name = graphicResourceName(graphic, "draw:fill-gradient-name");
    const resource = resources.gradients.find((candidate) => candidate.name === name);
    if (resource === undefined) throw missingGraphicResource(name, "draw:gradient");
    return gradientFill(resource);
  }
  if (token === "hatch") {
    const name = graphicResourceName(graphic, "draw:fill-hatch-name");
    const resource = resources.hatches.find((candidate) => candidate.name === name);
    if (resource === undefined) throw missingGraphicResource(name, "draw:hatch");
    return hatchFill(resource);
  }
  if (token === "bitmap") {
    const name = graphicResourceName(graphic, "draw:fill-image-name");
    const resource = resources.fillImages.find((candidate) => candidate.name === name);
    if (resource === undefined) throw missingGraphicResource(name, "draw:fill-image");
    const href = resource.attributes["xlink:href"];
    if (!href) throw missingGraphicResource(name, "draw:fill-image");
    const repeat = graphic.fillRepeat;
    if (repeat === "no-repeat")
      throw unsupportedGraphicStyle("style:repeat", "unsupported-bitmap-repeat");
    if (href.startsWith("#"))
      return { type: "blip", noEmbed: true, sourceReferenceId: href.slice(1) };
    if (isExternalReference(href)) return { type: "blip", linkedUrl: href };
    if (resource.data === undefined) throw missingGraphicResource(name, "draw:fill-image");
    return {
      type: "blip",
      data: resource.data,
      imageType: imageTypeFromHref(href),
      fileName: normalizePackagePath(href),
      ...(repeat === "repeat" ? { stretch: false } : {}),
    };
  }
  throw unsupportedGraphicStyle("draw:fill", "unsupported-odf-fill");
}

function projectGraphicOutline(graphic: GraphicStyle): OutlineOptions | undefined {
  if (graphic.stroke === "none") return { type: "noFill" };
  if (!graphic.strokeColor && !graphic.strokeWidth) return undefined;
  return {
    ...(graphic.strokeWidth ? { width: lengthToEmu(graphic.strokeWidth) ?? 0 } : {}),
    ...(graphic.strokeColor ? { type: "solidFill", color: graphic.strokeColor } : {}),
  };
}

function gradientFill(resource: GraphicStyleResource): FillOptions {
  const style = resource.attributes["draw:style"];
  const start = resource.attributes["draw:start-color"]?.replace(/^#/, "") ?? "000000";
  const end = resource.attributes["draw:end-color"]?.replace(/^#/, "") ?? "000000";
  const angle = Number(resource.attributes["draw:angle"] ?? 0);
  if (style === "linear")
    return {
      type: "gradient",
      angle,
      stops: [
        { position: 0, color: start },
        { position: 100, color: end },
      ],
    };
  if (style === "axial")
    return {
      type: "gradient",
      angle,
      stops: [
        { position: 0, color: start },
        { position: 50, color: end },
        { position: 100, color: start },
      ],
    };
  if (style === "radial" || style === "ellipsoid" || style === "rectangular")
    return {
      type: "gradient",
      path: style === "rectangular" ? "rect" : style === "ellipsoid" ? "shape" : "circle",
      stops: [
        { position: 0, color: start },
        { position: 100, color: end },
      ],
    };
  throw unsupportedGraphicStyle("draw:style", "unsupported-gradient");
}

function hatchFill(resource: GraphicStyleResource): FillOptions {
  const style = resource.attributes["draw:style"];
  const rotation = Number(resource.attributes["draw:rotation"] ?? 0);
  const pattern =
    style === "double"
      ? "cross"
      : style === "triple"
        ? "diagonalCross"
        : rotation === 450 || rotation === 1350
          ? "downDiagonal"
          : rotation === 900 || rotation === 2700
            ? "vertical"
            : "horizontal";
  return {
    type: "pattern",
    pattern,
    foregroundColor: resource.attributes["draw:color"]?.replace(/^#/, "") ?? "000000",
    backgroundColor: resource.attributes["draw:background-color"]?.replace(/^#/, "") ?? "FFFFFF",
  };
}

function fillResourceXml(
  fill: FillOptions | null | undefined,
  resources: GraphicStyleResources,
): { type: "gradient" | "hatch" | "bitmap"; attributes: Record<string, string> } | undefined {
  if (!fill || typeof fill === "string" || fill.type === "solid" || fill.type === "none")
    return undefined;
  if (fill.type === "gradient") {
    const gradient = "options" in fill ? fill.options : fill;
    if ((gradient.stops?.length ?? 0) > 2)
      throw unsupportedGraphicStyle("fill.stops", "odf-gradient-supports-two-stops");
    const first = gradient.stops?.[0];
    const last = gradient.stops?.at(-1);
    const shade =
      "shade" in gradient
        ? gradient.shade
        : "angle" in gradient && gradient.angle !== undefined
          ? { angle: gradient.angle }
          : undefined;
    const style =
      shade && "path" in shade
        ? shade.path === "rect"
          ? "rectangular"
          : shade.path === "shape"
            ? "ellipsoid"
            : "radial"
        : "linear";
    const resource = addGraphicResource(resources.gradients, "gradient", {
      "draw:style": style,
      ...(shade && "angle" in shade && shade.angle !== undefined
        ? { "draw:angle": String(shade.angle) }
        : {}),
      ...(first ? { "draw:start-color": `#${hexColorValue(first.color) ?? "000000"}` } : {}),
      ...(last ? { "draw:end-color": `#${hexColorValue(last.color) ?? "000000"}` } : {}),
    });
    return { type: "gradient", attributes: { "draw:fill-gradient-name": resource.name } };
  }
  if (fill.type === "pattern") {
    const resource = addGraphicResource(resources.hatches, "hatch", {
      "draw:style":
        fill.pattern === "cross" || fill.pattern === "diagonalCross" ? "double" : "single",
      "draw:rotation":
        fill.pattern === "vertical" ? "900" : fill.pattern === "downDiagonal" ? "450" : "0",
      ...(fill.foregroundColor
        ? { "draw:color": `#${hexColorValue(fill.foregroundColor) ?? "000000"}` }
        : {}),
      ...(fill.backgroundColor
        ? { "draw:background-color": `#${hexColorValue(fill.backgroundColor) ?? "FFFFFF"}` }
        : {}),
    });
    return { type: "hatch", attributes: { "draw:fill-hatch-name": resource.name } };
  }
  if (fill.type !== "blip") throw unsupportedGraphicStyle("fill.type", "unsupported-odf-fill");
  if (fill.noEmbed || fill.sourceReferenceId) {
    const href = fill.sourceReferenceId ? `#${fill.sourceReferenceId}` : undefined;
    if (!href) throw unsupportedGraphicStyle("fill.noEmbed", "missing-bitmap-reference");
    const resource = addGraphicResource(resources.fillImages, "fill-image", {
      "xlink:href": href,
    });
    return { type: "bitmap", attributes: { "draw:fill-image-name": resource.name } };
  }
  if (fill.linkedUrl) {
    const resource = addGraphicResource(resources.fillImages, "fill-image", {
      "xlink:href": fill.linkedUrl,
    });
    return { type: "bitmap", attributes: { "draw:fill-image-name": resource.name } };
  }
  if (!fill.data || !fill.imageType)
    throw unsupportedGraphicStyle("fill.data", "missing-bitmap-data");
  const fileName =
    fill.fileName ?? `Pictures/graphic-fill-${resources.fillImages.length + 1}.${fill.imageType}`;
  const resource = addFillImageResource(
    resources.fillImages,
    { "xlink:href": fileName },
    fill.data,
  );
  return {
    type: "bitmap",
    attributes: {
      "draw:fill-image-name": resource.name,
      ...(fill.stretch === false ? { "style:repeat": "repeat" } : {}),
    },
  };
}

function addGraphicResource(
  resources: GraphicStyleResource[],
  prefix: string,
  attributes: Record<string, string>,
): GraphicStyleResource {
  const existing = resources.find(
    (resource) => JSON.stringify(resource.attributes) === JSON.stringify(attributes),
  );
  if (existing) return existing;
  const resource = { name: `${prefix}-${resources.length + 1}`, attributes };
  resources.push(resource);
  return resource;
}

function addFillImageResource(
  resources: GraphicFillImageResource[],
  attributes: Record<string, string>,
  data: DataType,
): GraphicFillImageResource {
  const bytes = toUint8Array(data);
  const existing = resources.find(
    (resource) =>
      JSON.stringify(resource.attributes) === JSON.stringify(attributes) &&
      byteEqual(resource.data, bytes),
  );
  if (existing) return existing;
  const resource: GraphicFillImageResource = {
    name: `fill-image-${resources.length + 1}`,
    attributes,
    data: bytes,
  };
  resources.push(resource);
  return resource;
}

function byteEqual(left: Uint8Array | undefined, right: Uint8Array | undefined): boolean {
  if (left === undefined || right === undefined) return left === right;
  return (
    left.byteLength === right.byteLength && left.every((value, index) => value === right[index])
  );
}

function graphicResourceName(graphic: GraphicStyle, attribute: string): string | undefined {
  if (attribute === "draw:fill-gradient-name") return graphic.fillGradientName;
  if (attribute === "draw:fill-hatch-name") return graphic.fillHatchName;
  if (attribute === "draw:fill-image-name") return graphic.fillImageName;
  return undefined;
}

function missingGraphicResource(name: string | undefined, elementName: string): OdfSchemaError {
  return unsupportedGraphicStyle(
    elementName,
    name === undefined ? "missing-resource-name" : "missing-resource",
  );
}

function unsupportedGraphicStyle(elementName: string, reason: string): OdfSchemaError {
  const path =
    "/office:document-content/office:automatic-styles/style:style/style:graphic-properties";
  return new OdfSchemaError(
    `ODF graphic style ${elementName}: ${reason}`,
    "content.xml",
    path,
    elementName,
    reason,
  );
}

function isExternalReference(value: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("//") || value.startsWith("../");
}

function normalizePackagePath(path: string): string {
  return path
    .split("/")
    .filter((part) => part !== "" && part !== ".")
    .join("/");
}

function imageTypeFromHref(
  href: string,
): "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf" {
  const extension = href.split(/[?#]/u, 1)[0]?.split(".").pop()?.toLowerCase();
  const value = extension === "jpeg" ? "jpg" : extension;
  if (
    value === "png" ||
    value === "jpg" ||
    value === "gif" ||
    value === "bmp" ||
    value === "tif" ||
    value === "ico" ||
    value === "emf" ||
    value === "wmf"
  )
    return value;
  return "png";
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
