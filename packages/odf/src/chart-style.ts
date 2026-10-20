import type {
  ChartShapePropertiesOptions,
  DashStop,
  FillOptions,
  ChartStyleResource,
  ChartStyleResources,
  DrawingStyleResource,
  FillImageResource,
  OutlineOptions,
} from "@office-open/core";
import type { DataType } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import type { Element } from "@office-open/xml";

import type { OdfPackageFiles } from "./container/package";
import { hexColorValue, pushShapeStyle } from "./drawing/graphic-style";
import { OdfSchemaError } from "./errors";
import {
  attributeString,
  childNamed,
  childrenNamed,
  isExternalOdfUrl,
  lengthToEmu,
  xmlElement,
} from "./runtime/xml";

export interface ChartStyle {
  styleName?: string;
  graphic?: Record<string, string>;
  properties: Record<string, string>;
  part: string;
  containerPath: string;
}

const chartStyleNameCaches = new WeakMap<ChartStyleResources, Map<string, string>>();
const drawingResourceKeys = new WeakMap<ChartStyleResources, Map<string, string[]>>();

function chartNameCache(resources: ChartStyleResources): Map<string, string> {
  let cache = chartStyleNameCaches.get(resources);
  if (cache === undefined) {
    cache = new Map();
    chartStyleNameCaches.set(resources, cache);
  }
  return cache;
}

function resourceKeyCache(resources: ChartStyleResources): Map<string, string[]> {
  let cache = drawingResourceKeys.get(resources);
  if (cache === undefined) {
    cache = new Map();
    drawingResourceKeys.set(resources, cache);
  }
  return cache;
}

function uniqueStyleName(
  prefix: string,
  resources: ChartStyleResources,
  xmlStyles: readonly string[],
): string {
  const names = new Set([...resources.styles.map((style) => style.name), ...xmlStyles]);
  let index = resources.styles.length + xmlStyles.length + 1;
  while (names.has(`${prefix}-${index}`)) index += 1;
  return `${prefix}-${index}`;
}

function uniqueResourceName(resources: readonly FillImageResource[], prefix: string): string {
  const names = new Set(resources.map((resource) => resource.name));
  let index = resources.length + 1;
  while (names.has(`${prefix}-${index}`)) index += 1;
  return `${prefix}-${index}`;
}

function resourceFingerprint(data: DataType | undefined): string {
  if (data === undefined) return "";
  const bytes = toUint8Array(data);
  let hash = 2166136261;
  for (const byte of bytes) {
    hash ^= byte;
    hash = Math.imul(hash, 16777619);
  }
  return `${bytes.byteLength}:${(hash >>> 0).toString(16)}`;
}

function byteEqual(left: DataType | undefined, right: DataType | undefined): boolean {
  if (left === right) return true;
  if (left === undefined || right === undefined) return false;
  const leftBytes = toUint8Array(left);
  const rightBytes = toUint8Array(right);
  return (
    leftBytes.byteLength === rightBytes.byteLength &&
    leftBytes.every((byte, index) => byte === rightBytes[index])
  );
}

function dedupeResource<T extends DrawingStyleResource>(
  collection: T[],
  resources: ChartStyleResources,
  cachePrefix: string,
  attributes: Record<string, string>,
  create: (name: string) => T,
  data?: DataType,
): T {
  const attributesKey = JSON.stringify(attributes);
  const key = `${cachePrefix}:${attributesKey}:${resourceFingerprint(data)}`;
  const cache = resourceKeyCache(resources);
  for (const name of cache.get(key) ?? []) {
    const existing = collection.find((resource) => resource.name === name);
    if (
      existing &&
      JSON.stringify(existing.attributes) === attributesKey &&
      byteEqual((existing as FillImageResource).data, data)
    )
      return existing;
  }
  const names = new Set(collection.map((resource) => resource.name));
  let index = collection.length + 1;
  while (names.has(`${cachePrefix}-${index}`)) index += 1;
  const resource = create(`${cachePrefix}-${index}`);
  collection.push(resource);
  cache.set(key, [...(cache.get(key) ?? []), resource.name]);
  return resource;
}

export function parseChartStyleContext(
  roots: readonly { part: string; root: Element }[],
  binaries: Record<string, Uint8Array>,
  basePath: string,
): { resources: ChartStyleResources; styles: Map<string, ChartStyle> } {
  const resources = emptyResources();
  const styles = new Map<string, ChartStyle>();
  const namedStyles = new Map<string, ChartStyleResource>();

  for (const { part, root } of roots) {
    for (const containerName of ["office:automatic-styles", "office:styles"]) {
      const container = childNamed(root, containerName);
      for (const style of childrenNamed(container, "style:style")) {
        const name = attributeString(style, "style:name");
        if (!name || namedStyles.has(name)) continue;
        const resource = styleResource(style, part);
        namedStyles.set(name, resource);
        resources.styles.push(resource);
        const graphic = childNamed(style, "style:graphic-properties");
        const chart = childNamed(style, "style:chart-properties");
        styles.set(name, {
          styleName: name,
          graphic: elementAttributes(graphic),
          properties: elementAttributes(chart),
          part,
          containerPath: `/office:document-${part === "styles.xml" ? "styles" : "content"}/${containerName}/style:style[@style:name="${name}"]`,
        });
      }
    }
  }

  for (const { root } of roots) {
    const officeStyles = childNamed(root, "office:styles");
    collectResource(officeStyles, "draw:gradient", "draw:name", resources.gradients);
    collectResource(officeStyles, "draw:hatch", "draw:name", resources.hatches);
    collectResource(officeStyles, "draw:fill-image", "draw:name", resources.fillImages);
    collectResource(officeStyles, "draw:stroke-dash", "draw:name", resources.strokeDashes);
    for (const image of resources.fillImages) {
      if (image.data !== undefined || image.attributes["xlink:href"] === undefined) continue;
      const href = image.attributes["xlink:href"];
      if (isExternalOdfUrl(href)) continue;
      const path = resolvePackagePath(basePath, href);
      const data = binaries[path] ?? binaries[encodeURI(path)] ?? binaries[decodeURI(path)];
      if (data) image.data = data;
    }
  }
  return { resources, styles };
}

export function projectChartStyle(
  style: ChartStyle,
  resources: ChartStyleResources,
  basePath: string,
): ChartShapePropertiesOptions | undefined {
  const graphic = style.graphic ?? {};
  const fillToken = graphic["draw:fill"];
  const strokeToken = graphic["draw:stroke"];
  const fill = projectFill(style, resources, basePath);
  const outline = projectOutline(style, resources);
  if (fillToken === undefined && strokeToken === undefined) return undefined;
  return {
    ...(fill || outline ? { fill, outline } : {}),
    styleName: style.styleName,
  };
}

export function pushChartStyle(
  shapeProperties: ChartShapePropertiesOptions | undefined,
  resources: ChartStyleResources,
  styles: string[],
  fallbackName: string,
  chartProperties: Record<string, string> = {},
): string | undefined {
  if (shapeProperties?.styleName) {
    const style = resources.styles.find((item) => item.name === shapeProperties.styleName);
    if (!style)
      throw styleError(
        "content.xml",
        "/office:document-content/office:automatic-styles",
        shapeProperties.styleName,
        "unresolved-style-reference",
      );
    return style.name;
  }
  const resource = pushChartFillResource(shapeProperties?.fill, resources);
  const graphicName =
    resource?.name ??
    pushShapeStyle(shapeProperties?.fill, shapeProperties?.outline, styles, fallbackName);
  if (resource && Object.values(chartProperties).every((value) => value === undefined)) {
    const cache = chartNameCache(resources);
    const cacheKey = `chart-fill:${resource.name}`;
    const cachedName = cache.get(cacheKey);
    if (cachedName !== undefined) return cachedName;
    const name = uniqueStyleName("chart-fill", resources, styles);
    styles.push(
      xmlElement("style:style", { "style:name": name, "style:family": "chart" }, [
        resource.graphicXml,
      ]),
    );
    cache.set(cacheKey, name);
    return name;
  }
  if (Object.values(chartProperties).every((value) => value === undefined)) return graphicName;
  const name = uniqueStyleName(`ch-${fallbackName}`, resources, styles);
  const graphicXml = graphicName
    ? (styles
        .find((style) => style.includes(`style:name="${graphicName}"`))
        ?.replace(/^<style:style[^>]*>|<\/style:style>$/g, "") ?? "")
    : "";
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "chart" }, [
      xmlElement("style:chart-properties", chartProperties),
      ...(graphicXml ? [graphicXml] : []),
    ]),
  );
  return name;
}

function pushChartFillResource(
  fill: FillOptions | null | undefined,
  resources: ChartStyleResources,
): { name: string; graphicXml: string } | undefined {
  if (!fill || typeof fill === "string" || fill.type === "solid" || fill.type === "none")
    return undefined;
  if (fill.type === "gradient") {
    const gradient = "options" in fill ? fill.options : fill;
    const stops = gradient.stops;
    const angle = "angle" in gradient ? gradient.angle : undefined;
    const path = "path" in gradient ? gradient.path : undefined;
    const gradientAttributes = {
      ...(path
        ? {
            "draw:style":
              path === "rect" ? "rectangular" : path === "circle" ? "radial" : "ellipsoid",
          }
        : { "draw:style": "linear" }),
      ...(angle !== undefined ? { "draw:angle": String(angle) } : {}),
      ...(stops?.[0]
        ? { "draw:start-color": `#${hexColorValue(stops[0].color) ?? "000000"}` }
        : {}),
      ...(stops?.at(-1)
        ? { "draw:end-color": `#${hexColorValue(stops.at(-1)!.color) ?? "000000"}` }
        : {}),
    };
    const resource = dedupeResource(
      resources.gradients,
      resources,
      "gradient",
      gradientAttributes,
      (name) => ({ name, attributes: gradientAttributes }),
    );
    const name = resource.name;
    return {
      name,
      graphicXml: xmlElement("style:graphic-properties", {
        "draw:fill": "gradient",
        "draw:fill-gradient-name": name,
      }),
    };
  }
  if (fill.type === "pattern") {
    const hatchAttributes = {
      "draw:style":
        fill.pattern === "cross" || fill.pattern === "diagonalCross" ? "double" : "single",
      "draw:rotation":
        fill.pattern === "vertical" ? "900" : fill.pattern === "downDiagonal" ? "450" : "0",
      ...(fill.foregroundColor
        ? { "draw:color": `#${hexColorValue(fill.foregroundColor) ?? "000000"}` }
        : {}),
    };
    const resource = dedupeResource(
      resources.hatches,
      resources,
      "hatch",
      hatchAttributes,
      (name) => ({ name, attributes: hatchAttributes }),
    );
    const name = resource.name;
    return {
      name,
      graphicXml: xmlElement("style:graphic-properties", {
        "draw:fill": "hatch",
        "draw:fill-hatch-name": name,
      }),
    };
  }
  if (fill.type !== "blip")
    throw styleError(
      "content.xml",
      "/office:document-content",
      fill.type,
      "unsupported-chart-fill",
    );
  const href: string =
    fill.linkedUrl ??
    (fill.sourceReferenceId
      ? `#${fill.sourceReferenceId}`
      : `media/${fill.fileName ?? `${uniqueResourceName(resources.fillImages, "fill-image")}.png`}`);
  const resource = dedupeResource(
    resources.fillImages,
    resources,
    "fill-image",
    { "xlink:href": href },
    (name) => ({
      name,
      attributes: { "xlink:href": href },
      ...(fill.data === undefined ? {} : { data: fill.data }),
    }),
    fill.data,
  );
  const name = resource.name;
  return {
    name,
    graphicXml: xmlElement("style:graphic-properties", {
      "draw:fill": "bitmap",
      "draw:fill-image-name": name,
      ...(fill.stretch === false ? { "style:repeat": "repeat" } : {}),
    }),
  };
}

export function chartStylesXml(resources: ChartStyleResources): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:styles>${[
    ...resources.styles.map((style) => styleResourceXml("style:style", style)),
    ...resources.gradients.map((resource) => styleResourceXml("draw:gradient", resource)),
    ...resources.hatches.map((resource) => styleResourceXml("draw:hatch", resource)),
    ...resources.fillImages.map((resource) => styleResourceXml("draw:fill-image", resource)),
    ...resources.strokeDashes.map((resource) => styleResourceXml("draw:stroke-dash", resource)),
  ].join("")}</office:styles></office:document-styles>`;
}

export function chartMediaFiles(basePath: string, resources: ChartStyleResources): OdfPackageFiles {
  const files: OdfPackageFiles = {};
  for (const image of resources.fillImages) {
    if (image.data === undefined) continue;
    const href = image.attributes["xlink:href"] ?? defaultImageHref(basePath, image.name);
    image.attributes["xlink:href"] ??= href;
    if (!isExternalOdfUrl(href))
      files[resolvePackagePath(basePath, href)] = toUint8Array(image.data);
  }
  return files;
}

function emptyResources(): ChartStyleResources {
  return {
    styles: [],
    gradients: [],
    hatches: [],
    fillImages: [],
    strokeDashes: [],
    dr3dLights: [],
  };
}

function elementAttributes(element: Element | undefined): Record<string, string> {
  return Object.fromEntries(
    Object.entries(element?.attributes ?? {}).map(([key, value]) => [key, String(value)]),
  );
}

function styleResource(element: Element, part: string): ChartStyleResource {
  const name = attributeString(element, "style:name");
  if (!name)
    throw styleError(
      part,
      "/office:document-content/office:automatic-styles/style:style",
      "style:name",
      "invalid-style-resource",
    );
  return {
    name,
    displayName: attributeString(element, "style:display-name"),
    attributes: withoutAttributes(element, ["style:name", "style:display-name"]),
    properties: (element.elements ?? [])
      .filter((child) => child.type === "element" && child.name?.startsWith("style:"))
      .map((child) => ({ name: child.name!, attributes: elementAttributes(child as Element) })),
  };
}

function collectResource(
  container: Element | undefined,
  elementName: string,
  nameAttribute: string,
  output: DrawingStyleResource[],
): void {
  for (const element of childrenNamed(container, elementName)) {
    const name = attributeString(element, nameAttribute);
    if (!name)
      throw styleError(
        "styles.xml",
        `/office:document-styles/office:styles/${elementName}`,
        nameAttribute,
        "invalid-style-resource",
      );
    if (output.some((resource) => resource.name === name)) continue;
    output.push({
      name,
      displayName: attributeString(element, "draw:display-name"),
      attributes: withoutAttributes(element, [nameAttribute, "draw:display-name"]),
    });
  }
}

function withoutAttributes(element: Element, excluded: readonly string[]): Record<string, string> {
  return Object.fromEntries(
    Object.entries(element.attributes ?? {})
      .filter(([key]) => !excluded.includes(key))
      .map(([key, value]) => [key, String(value)]),
  );
}

function projectFill(
  style: ChartStyle,
  resources: ChartStyleResources,
  basePath: string,
): FillOptions | undefined {
  const graphic = style.graphic ?? {};
  const token = graphic["draw:fill"];
  if (token === undefined) return undefined;
  if (token === "none") return { type: "none" };
  if (token === "solid") {
    const color = stripColor(graphic["draw:fill-color"]) ?? "000000";
    return { type: "solid", color };
  }
  if (token === "gradient") {
    const name = graphic["draw:fill-gradient-name"];
    const resource = name === undefined ? undefined : findResource(resources.gradients, name);
    if (resource === undefined) return undefined;
    return gradientFill(resource);
  }
  if (token === "hatch") {
    const name = graphic["draw:fill-hatch-name"];
    const resource = name === undefined ? undefined : findResource(resources.hatches, name);
    if (resource === undefined) return undefined;
    return hatchFill(resource);
  }
  if (token === "bitmap") return bitmapFill(style, resources, basePath);
  throw projectionError(style, "draw:fill", "unsupported-odf-fill");
}

function hatchFill(resource: DrawingStyleResource): FillOptions {
  const style = resource.attributes["draw:style"];
  const rotation = ((Number(resource.attributes["draw:rotation"] ?? 0) % 3600) + 3600) % 3600;
  const foreground = stripColor(resource.attributes["draw:color"]) ?? "000000";
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
  return { type: "pattern", pattern, foregroundColor: foreground, backgroundColor: "FFFFFF" };
}

function bitmapFill(
  style: ChartStyle,
  resources: ChartStyleResources,
  basePath: string,
): FillOptions | undefined {
  const graphic = style.graphic ?? {};
  const name = graphic["draw:fill-image-name"];
  const resource = resources.fillImages.find((item) => item.name === name);
  if (resource === undefined) return undefined;
  const href = resource.attributes["xlink:href"];
  if (href?.startsWith("#")) {
    return { type: "blip", noEmbed: true, sourceReferenceId: href.slice(1) };
  }
  const repeat = graphic["style:repeat"] as "no-repeat" | "repeat" | "stretch" | undefined;
  const overlay = {
    type: "bitmap" as const,
    name,
    ...(repeat ? { repeat } : {}),
    ...(graphic["draw:fill-image-width"] ? { width: graphic["draw:fill-image-width"] } : {}),
    ...(graphic["draw:fill-image-height"] ? { height: graphic["draw:fill-image-height"] } : {}),
    ...(graphic["draw:fill-image-ref-point"]
      ? { refPoint: graphic["draw:fill-image-ref-point"] }
      : {}),
    ...(graphic["draw:fill-image-ref-point-x"]
      ? { refPointX: graphic["draw:fill-image-ref-point-x"] }
      : {}),
    ...(graphic["draw:fill-image-ref-point-y"]
      ? { refPointY: graphic["draw:fill-image-ref-point-y"] }
      : {}),
    ...(graphic["draw:tile-repeat-offset"]
      ? { tileRepeatOffset: graphic["draw:tile-repeat-offset"] }
      : {}),
    ...(graphic["draw:opacity-name"] ? { opacityName: graphic["draw:opacity-name"] } : {}),
  };
  if (!href) throw projectionError(style, "draw:fill-image-name", "missing-fill-image-part");
  const data = resource.data;
  const linkedUrl = isExternalOdfUrl(href) ? href : undefined;
  if (!data && !linkedUrl)
    throw projectionError(style, "draw:fill-image-name", "missing-fill-image-part");
  const unsupportedBitmap =
    repeat === "no-repeat" ||
    overlay.tileRepeatOffset !== undefined ||
    overlay.opacityName !== undefined;
  const imageType = imageTypeFromHref(href);
  const tile = unsupportedBitmap || repeat !== "repeat" ? undefined : bitmapTile(graphic, resource);
  const value: FillOptions | undefined = unsupportedBitmap
    ? undefined
    : {
        type: "blip",
        ...(data
          ? { data, imageType, fileName: resolvePackagePath(basePath, href) }
          : { linkedUrl }),
        ...(tile ? { stretch: false, tile } : {}),
      };
  return value;
}

function bitmapTile(graphic: Record<string, string>, resource: FillImageResource) {
  const sx = tileScale(graphic["draw:fill-image-width"], resource.attributes["svg:width"]);
  const sy = tileScale(graphic["draw:fill-image-height"], resource.attributes["svg:height"]);
  if (sx === undefined || sy === undefined) return undefined;
  return {
    sx,
    sy,
    ...(graphic["draw:fill-image-ref-point"]
      ? { alignment: tileAlignment(graphic["draw:fill-image-ref-point"]) }
      : {}),
  };
}

function projectOutline(
  style: ChartStyle,
  resources: ChartStyleResources,
): OutlineOptions | undefined {
  const graphic = style.graphic ?? {};
  if (graphic["draw:stroke"] === "none") return { type: "noFill" };
  if (graphic["draw:stroke"] !== "dash") {
    if (graphic["svg:stroke-width"] || graphic["svg:stroke-color"]) {
      return {
        width: lengthEmu(graphic["svg:stroke-width"], style, "svg:stroke-width") ?? 0,
        type: "solidFill",
        ...(graphic["svg:stroke-color"] ? { color: stripColor(graphic["svg:stroke-color"]) } : {}),
      };
    }
    return undefined;
  }
  const name = graphic["draw:stroke-dash"];
  const resource = findResource(resources.strokeDashes, name);
  if (!name || !resource)
    throw projectionError(style, "draw:stroke-dash", "unresolved-style-reference");
  const dash = customDash(resource, style);
  return {
    width:
      lengthEmu(graphic["svg:stroke-width"], style, "svg:stroke-width") ??
      (() => {
        throw projectionError(style, "svg:stroke-width", "missing-dash-basis");
      })(),
    type: "solidFill",
    ...(graphic["svg:stroke-color"] ? { color: stripColor(graphic["svg:stroke-color"]) } : {}),
    customDash: dash,
  };
}

function customDash(resource: DrawingStyleResource, style: ChartStyle): readonly DashStop[] {
  if (resource.attributes["draw:style"] === "round") return [];
  const width = lengthEmu(style.graphic?.["svg:stroke-width"], style, "svg:stroke-width");
  if (!width) throw projectionError(style, "svg:stroke-width", "missing-dash-basis");
  const stops: DashStop[] = [];
  const firstCount = Number(resource.attributes["draw:dots1"] ?? 1);
  const secondCount = Number(resource.attributes["draw:dots2"] ?? 0);
  const distance = dashLength(resource.attributes["draw:distance"], width, style);
  if (Number.isFinite(firstCount) && firstCount > 0)
    for (let index = 0; index < firstCount; index++)
      stops.push({
        d: dashLength(resource.attributes["draw:dots1-length"] ?? "100%", width, style),
        sp: distance,
      });
  if (Number.isFinite(secondCount) && secondCount > 0)
    for (let index = 0; index < secondCount; index++)
      stops.push({
        d: dashLength(resource.attributes["draw:dots2-length"] ?? "100%", width, style),
        sp: distance,
      });
  return stops;
}

function dashLength(value: string | undefined, widthEmu: number, style: ChartStyle): `${number}%` {
  const percentage = value?.endsWith("%") ? Number(value.slice(0, -1)) : undefined;
  const ratio =
    percentage !== undefined
      ? percentage / 100
      : value === undefined
        ? undefined
        : (lengthToEmu(value) ?? Number.NaN) / widthEmu;
  if (ratio === undefined || !Number.isFinite(ratio) || ratio <= 0)
    throw projectionError(style, "draw:stroke-dash", "unsupported-dash-shape");
  return `${Number((ratio * 100).toFixed(6))}%`;
}

function tileScale(
  value: string | undefined,
  intrinsicValue: string | undefined,
): number | undefined {
  if (value === undefined) return 100;
  if (value.endsWith("%")) {
    const percent = Number(value.slice(0, -1));
    if (!Number.isFinite(percent) || percent <= 0) return undefined;
    return percent;
  }
  const size = lengthToEmu(value);
  const intrinsic = intrinsicValue === undefined ? undefined : lengthToEmu(intrinsicValue);
  if (!size || !intrinsic) return undefined;
  return Number(((size / intrinsic) * 100).toFixed(6));
}

function lengthEmu(
  value: string | undefined,
  style: ChartStyle,
  property: string,
): number | undefined {
  if (value === undefined) return undefined;
  const emu = lengthToEmu(value);
  if (emu === undefined) throw projectionError(style, property, "invalid-odf-length");
  return emu;
}

function tileAlignment(
  value: string,
):
  | "topLeft"
  | "top"
  | "topRight"
  | "left"
  | "center"
  | "right"
  | "bottomLeft"
  | "bottom"
  | "bottomRight" {
  const mapped = value
    .split("-")
    .map((part, index) =>
      index === 0 ? part : `${part[0]!.toUpperCase()}${part.slice(1).toLowerCase()}`,
    )
    .join("") as
    | "topLeft"
    | "top"
    | "topRight"
    | "left"
    | "center"
    | "right"
    | "bottomLeft"
    | "bottom"
    | "bottomRight";
  return mapped;
}

function scaledColor(value: string | undefined, intensity: string | undefined): string {
  const color = stripColor(value) ?? "000000";
  const percent = intensity === undefined ? 100 : Number(intensity.replace(/%$/, ""));
  if (!Number.isFinite(percent) || percent < 0 || percent > 100) return color;
  const factor = percent / 100;
  const channels = [0, 2, 4].map((index) =>
    Math.round(parseInt(color.slice(index, index + 2), 16) * factor),
  );
  return channels
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function stripColor(value: string | undefined): string | undefined {
  return value?.replace(/^#/, "").toUpperCase();
}

function imageTypeFromHref(
  href: string,
): "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf" {
  const extension = /\.([a-z0-9]+)(?:[?#].*)?$/i.exec(href)?.[1]?.toLowerCase();
  const mapped = extension === "jpeg" ? "jpg" : extension;
  if (mapped && ["png", "jpg", "gif", "bmp", "tif", "ico", "emf", "wmf"].includes(mapped))
    return mapped as "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
  return "png";
}

function defaultImageHref(basePath: string, name: string): string {
  return `${basePath ? `${basePath}media/` : "media/"}${name}.png`;
}

function resolvePackagePath(basePath: string, href: string): string {
  if (href.startsWith("/")) return normalizePackagePath(href.slice(1));
  return normalizePackagePath(`${basePath}${href}`);
}

function normalizePackagePath(path: string): string {
  const decoded = decodeURIComponent(path);
  const pieces: string[] = [];
  for (const piece of decoded.split("/")) {
    if (!piece || piece === ".") continue;
    if (piece === "..") pieces.pop();
    else pieces.push(piece);
  }
  return `${pieces.join("/")}`;
}

function gradientFill(resource: DrawingStyleResource): FillOptions {
  const kind = resource.attributes["draw:style"];
  if (kind === "square") return { type: "gradient", stops: [] };
  const start = scaledColor(
    resource.attributes["draw:start-color"],
    resource.attributes["draw:start-intensity"],
  );
  const end = scaledColor(
    resource.attributes["draw:end-color"],
    resource.attributes["draw:end-intensity"],
  );
  const angle = Number(resource.attributes["draw:angle"] ?? 0);
  if (kind === "linear")
    return {
      type: "gradient",
      angle,
      stops: [
        { position: 0, color: start },
        { position: 100, color: end },
      ],
    };
  if (kind === "axial")
    return {
      type: "gradient",
      angle,
      stops: [
        { position: 0, color: start },
        { position: 50, color: end },
        { position: 100, color: start },
      ],
    };
  if (kind === "radial" || kind === "ellipsoid" || kind === "rectangular")
    return {
      type: "gradient",
      path: kind === "rectangular" ? "rect" : kind === "ellipsoid" ? "shape" : "circle",
      stops: [
        { position: 0, color: start },
        { position: 100, color: end },
      ],
    };
  throw styleError(
    "styles.xml",
    `/office:document-styles/office:styles/draw:gradient[@draw:name="${resource.name}"]`,
    resource.name,
    "unsupported-gradient-for-drawingml",
  );
}

function styleResourceXml(
  name: string,
  resource: DrawingStyleResource | ChartStyleResource,
): string {
  const nameAttribute = name === "style:style" ? "style:name" : "draw:name";
  const displayAttribute = name === "style:style" ? "style:display-name" : "draw:display-name";
  const properties =
    "properties" in resource
      ? resource.properties.map((property) => xmlElement(property.name, property.attributes))
      : [];
  const attributes = {
    [nameAttribute]: resource.name,
    [displayAttribute]: resource.displayName,
    ...resource.attributes,
  };
  if (name === "draw:fill-image") {
    attributes["xlink:type"] ??= "simple";
    attributes["xlink:show"] ??= "embed";
    attributes["xlink:actuate"] ??= "onLoad";
  }
  return xmlElement(name, attributes, properties);
}

function findResource(resources: readonly DrawingStyleResource[], name: string | undefined) {
  return resources.find((resource) => resource.name === name);
}

function projectionError(style: ChartStyle, property: string, reason: string): OdfSchemaError {
  return styleError(
    style.part,
    `${style.containerPath}/style:graphic-properties/@${property}`,
    style.styleName ?? "",
    reason,
  );
}

function styleError(part: string, path: string, name: string, reason: string): OdfSchemaError {
  return new OdfSchemaError(`${part}: ${path}: ${name}: ${reason}`, part, path, name, reason);
}
