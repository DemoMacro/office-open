import { toUint8Array, type ChartSpaceOptions } from "@office-open/core";
import type { FillOptions, OutlineOptions } from "@office-open/core/drawing";
import type {
  ChartOptions,
  GroupChildMediaData,
  GroupOptions,
  ParagraphChild,
  PictureOptions,
  SectionChild,
  ShapeOptions,
} from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  graphicFill,
  graphicOutline,
  isExternalOdfUrl,
  lengthToEmu,
  PRESET_GEOMETRY_DOCX,
  presetGeometryOdf,
  pushShapeStyle,
  resolveEmbeddedChart,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

type WpgChildMediaData = Extract<GroupChildMediaData, { type: "wps" | "wpg" }>;

import { unknownOdtElement, type ParseContext } from "./body";
import type {
  OdtEmbeddedObjectFrameOptions,
  OdtEmbeddedObjectOptions,
  OdtFrameAnchorType,
} from "./embedded-objects";
import { OdtParseError } from "./error";

export type PictureType = NonNullable<PictureOptions["type"]>;

export type RasterPictureType = Exclude<PictureType, "svg">;

type DrawingTransformation = ChartOptions["transformation"];

interface ParsedImage {
  type: PictureType;
  data: Uint8Array;
}

function transformationWidth(transformation: DrawingTransformation): number | string {
  return typeof transformation.width === "number"
    ? emuToLength(transformation.width)
    : transformation.width;
}

function transformationHeight(transformation: DrawingTransformation): number | string {
  return typeof transformation.height === "number"
    ? emuToLength(transformation.height)
    : transformation.height;
}

function transformationOffset(
  offset: DrawingTransformation["offset"],
  axis: "x" | "y",
): number | string | undefined {
  const value = axis === "x" ? offset?.x : offset?.y;
  return value === undefined || value === 0 ? undefined : emuToLength(value);
}

/** Binary image collected during generation — emitted as a Pictures/ entry. */
export interface OdtImage {
  path: string;
  data: Uint8Array;
}

/** Embedded chart subdocument collected during generation. */
export interface OdtChart {
  path: string;
  chart: ChartSpaceOptions;
}

export function isRasterPictureType(value: string): value is RasterPictureType {
  return ["png", "jpg", "gif", "bmp", "tif", "ico", "emf", "wmf"].includes(value);
}

export function isPictureType(value: string): value is PictureType {
  return value === "svg" || isRasterPictureType(value);
}

/** Inline chart renders as a draw:frame + draw:object pointing at the subdocument. */
export function chartFrameXml(chart: ChartOptions, charts: OdtChart[]): string {
  const path = `Object ${charts.length + 1}`;
  charts.push({ path, chart });
  return xmlElement(
    "draw:frame",
    {
      "text:anchor-type": "as-char",
      "svg:width": transformationWidth(chart.transformation),
      "svg:height": transformationHeight(chart.transformation),
    },
    [xmlElement("draw:object", { "xlink:href": `./${path}`, "xlink:type": "simple" })],
  );
}

/** Inline shape renders as a positioned draw:custom-shape with preset geometry. */
export function wpsShapeFrameXml(
  shape: ShapeOptions,
  styles: string[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const geometry = typeof shape.geometry === "string" ? { preset: shape.geometry } : shape.geometry;
  const offset = shape.transformation.offset;
  const shapeId = shape.nonVisualProperties?.id;
  const object =
    shapeId === undefined
      ? undefined
      : embeddedObjects.find((candidate) =>
          candidate.frames?.some((frame) => frame.shapeId === shapeId),
        );
  if (shape.controlId !== undefined) {
    return xmlElement(
      "draw:control",
      {
        "draw:control": shape.controlId,
        "text:anchor-type": "as-char",
        "draw:style-name": addShapeStyle(shape, styles),
        "svg:x": transformationOffset(offset, "x"),
        "svg:y": transformationOffset(offset, "y"),
        "svg:width": transformationWidth(shape.transformation),
        "svg:height": transformationHeight(shape.transformation),
        "draw:name": shape.altText?.name,
      },
      [],
    );
  }
  if (object) {
    const frame = object.frames?.find((candidate) => candidate.shapeId === shapeId);
    if (!frame) throw malformedEmbeddedObject("embedded object frame reference is missing");
    return embeddedObjectFrameXml(shape, object, frame, styles);
  }
  return xmlElement(
    "draw:custom-shape",
    {
      "text:anchor-type": "as-char",
      "draw:style-name": addShapeStyle(shape, styles),
      "svg:x": transformationOffset(offset, "x"),
      "svg:y": transformationOffset(offset, "y"),
      "svg:width": transformationWidth(shape.transformation),
      "svg:height": transformationHeight(shape.transformation),
      "draw:name": shape.altText?.name,
    },
    [
      xmlElement(
        "draw:enhanced-geometry",
        { "draw:type": geometry?.preset ? presetGeometryOdf(geometry.preset) : undefined },
        [],
      ),
    ],
  );
}

/** Canonical wpsShape plus an object overlay restores draw:frame + draw:object. */
function embeddedObjectFrameXml(
  shape: ShapeOptions,
  object: OdtEmbeddedObjectOptions,
  frame: OdtEmbeddedObjectFrameOptions,
  styles: string[],
): string {
  const offset = shape.transformation.offset;
  const name = shape.nonVisualProperties?.name ?? shape.altText?.name;
  return xmlElement(
    "draw:frame",
    {
      "text:anchor-type": frame.anchorType,
      "draw:style-name": frame.styleName ?? addShapeStyle(shape, styles),
      "svg:x": lengthAttributeXml(transformationOffset(offset, "x")),
      "svg:y": lengthAttributeXml(transformationOffset(offset, "y")),
      "svg:width": lengthAttributeXml(transformationWidth(shape.transformation)),
      "svg:height": lengthAttributeXml(transformationHeight(shape.transformation)),
      "draw:name": name,
    },
    [xmlElement("draw:object", { "xlink:href": `./${object.path}`, "xlink:type": "simple" })],
  );
}

function lengthAttributeXml(value: number | string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return typeof value === "number" ? emuToLength(value) : value;
}

/** Shape fill and outline land in a reusable graphic style. */
export function addShapeStyle(shape: ShapeOptions, styles: string[]): string | undefined {
  return pushShapeStyle(shape.fill, shape.outline, styles);
}

/** Inline picture renders as a character-anchored draw:frame + draw:image. */
export function pictureFrameXml(picture: PictureOptions, images: OdtImage[]): string {
  const raster = picture.type === "svg" ? picture.fallback : picture;
  let path: string;
  if (raster.data === undefined) {
    if (!picture.sourceUrl) return "";
    path = picture.sourceUrl;
  } else {
    const data = toUint8Array(raster.data);
    path = `Pictures/picture${images.length + 1}.${raster.type}`;
    images.push({ path, data });
  }
  return xmlElement(
    "draw:frame",
    {
      "text:anchor-type": "as-char",
      "svg:width": transformationWidth(picture.transformation),
      "svg:height": transformationHeight(picture.transformation),
      "draw:name": picture.altText?.name,
    },
    [
      xmlElement("draw:image", {
        "xlink:href": path,
        ...(picture.sourceUrl && raster.data === undefined
          ? { "xlink:type": "simple", "xlink:show": "embed" }
          : {}),
      }),
    ],
  );
}

export function shapeBlock(shape: ShapeOptions): SectionChild {
  return { paragraph: { children: [{ wpsShape: shape }] } };
}

export function parseShape(element: Element): ShapeOptions {
  return parseShapeElement(element, undefined);
}

export function parseShapeElement(
  element: Element,
  context: ParseContext | undefined,
): ShapeOptions {
  const name = attributeString(element, "draw:name");
  const styleName = attributeString(element, "draw:style-name");
  const graphic = context?.graphicStyles.get(styleName ?? "");
  const x1 = lengthToEmu(attributeString(element, "svg:x1"));
  const y1 = lengthToEmu(attributeString(element, "svg:y1"));
  const x2 = lengthToEmu(attributeString(element, "svg:x2"));
  const y2 = lengthToEmu(attributeString(element, "svg:y2"));
  const hasEndpoints = x1 !== undefined || y1 !== undefined || x2 !== undefined || y2 !== undefined;
  const x = hasEndpoints
    ? Math.min(x1 ?? 0, x2 ?? 0)
    : (lengthToEmu(attributeString(element, "svg:x")) ?? 0);
  const y = hasEndpoints
    ? Math.min(y1 ?? 0, y2 ?? 0)
    : (lengthToEmu(attributeString(element, "svg:y")) ?? 0);
  const width = hasEndpoints
    ? Math.abs((x2 ?? 0) - (x1 ?? 0))
    : (lengthToEmu(attributeString(element, "svg:width")) ?? 0);
  const height = hasEndpoints
    ? Math.abs((y2 ?? 0) - (y1 ?? 0))
    : (lengthToEmu(attributeString(element, "svg:height")) ?? 0);
  const enhanced = childNamed(element, "draw:enhanced-geometry");
  const presetType = attributeString(enhanced, "draw:type");
  const preset = presetType ? PRESET_GEOMETRY_DOCX[presetType] : undefined;
  const geometry = shapeGeometry(element.name, preset);
  const fill = graphic ? graphicFill(graphic) : undefined;
  const outline = graphic ? graphicOutline(graphic) : undefined;
  const controlId = attributeString(element, "draw:control");
  return {
    children: [],
    ...(controlId === undefined ? {} : { controlId }),
    ...(geometry ? { geometry } : {}),
    ...(name ? { altText: { name } } : {}),
    transformation: {
      ...(x || y ? { offset: { ...(x ? { x } : {}), ...(y ? { y } : {}) } } : {}),
      width,
      height,
    },
    ...(fill ? { fill } : {}),
    ...(outline ? { outline } : {}),
  };
}

function shapeGeometry(
  name: string | undefined,
  preset: ShapeOptions["geometry"] | undefined,
): ShapeOptions["geometry"] | undefined {
  if (preset) return preset;
  if (name === "draw:rect" || name === "draw:frame" || name === "draw:control") return "rect";
  if (name === "draw:ellipse" || name === "draw:circle") return "ellipse";
  if (
    name === "draw:line" ||
    name === "draw:connector" ||
    name === "draw:measure" ||
    name === "draw:polyline"
  )
    return "line";
  return undefined;
}

export function parseShapeHyperlink(element: Element): ParagraphChild {
  return {
    hyperlink: {
      url: attributeString(element, "xlink:href") ?? "",
      children: (element.elements ?? [])
        .filter((child) => isShapeElement(child.name))
        .map((shape) => ({ wpsShape: parseShapeElement(shape, undefined) })),
    },
  };
}

export function isShapeElement(name: string | undefined): boolean {
  return (
    name === "draw:rect" ||
    name === "draw:control" ||
    name === "draw:ellipse" ||
    name === "draw:circle" ||
    name === "draw:line" ||
    name === "draw:connector" ||
    name === "draw:measure" ||
    name === "draw:path" ||
    name === "draw:polygon" ||
    name === "draw:polyline" ||
    name === "draw:custom-shape"
  );
}

/** draw:frame + draw:image maps back to an inline picture run. */
export function parsePictureFrame(frame: Element, context: ParseContext): ParagraphChild[] {
  const images: ParsedImage[] = childrenNamed(frame, "draw:image").flatMap((image) => {
    const href = attributeString(image, "xlink:href");
    const path = href?.replace(/^\//, "");
    const data = path ? context.binaries[path] : undefined;
    if (data === undefined) return [];
    context.usedBinaryPaths.add(path ?? "");
    const extension = path?.split(".").pop() ?? "png";
    const type = extension === "jpeg" ? "jpg" : extension;
    return isPictureType(type) ? [{ type, data }] : [];
  });
  const svg = images.find((image) => image.type === "svg");
  const raster = images.find(
    (image): image is ParsedImage & { type: RasterPictureType } =>
      image.type !== "svg" && isRasterPictureType(image.type),
  );
  if (!svg && !raster) {
    const href = attributeString(childNamed(frame, "draw:image"), "xlink:href");
    const packagePath = href?.replace(/^\//, "");
    const missingPackageImage =
      packagePath === undefined || context.binaries[packagePath] === undefined;
    if (href && (isExternalOdfUrl(href) || missingPackageImage))
      return [
        {
          picture: {
            type: pictureTypeFromUrl(href),
            sourceUrl: href,
            transformation: {
              width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
              height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
            },
          },
        },
      ];
    return [];
  }
  return [
    {
      picture: {
        ...(svg
          ? {
              type: "svg",
              data: svg.data,
              fallback: {
                type: raster?.type ?? "png",
                data: raster?.data ?? new Uint8Array(),
              },
            }
          : {
              type: raster?.type ?? "png",
              data: raster?.data ?? new Uint8Array(),
            }),
        altText: {
          name: attributeString(frame, "draw:name") ?? "",
          description: textOf(childNamed(frame, "svg:desc")) || undefined,
        },
        transformation: {
          width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
          height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
        },
      },
    },
  ];
}

function pictureTypeFromUrl(value: string): RasterPictureType {
  const extension = value.split(/[?#]/, 1)[0]?.split("/").pop()?.split(".").pop()?.toLowerCase();
  if (extension === "jpeg") return "jpg";
  return isRasterPictureType(extension ?? "") ? (extension as RasterPictureType) : "png";
}

/** draw:frame + draw:object resolves a typed embedded-object subdocument. */
export function parseChartFrame(frame: Element, context: ParseContext): ParagraphChild[] {
  const href = attributeString(childNamed(frame, "draw:object"), "xlink:href")
    ?.replace(/^\.\//, "")
    .replace(/^\//, "");
  const chart = href ? resolveEmbeddedChart(context.chartBodies, href) : undefined;
  if (!chart) {
    const mediaType = href ? context.objectMediaTypes.get(href) : undefined;
    if (!href) throw malformedEmbeddedObject("chart frame has no object reference", undefined);
    if (mediaType === undefined)
      throw malformedEmbeddedObject(
        "embedded object media type is not declared",
        href,
        "missing-package-part",
      );
    if (!context.embeddedObjects.has(href))
      throw malformedEmbeddedObject("embedded object subdocument is missing", href);
  }
  if (chart) {
    return [
      {
        chart: {
          ...chart,
          transformation: {
            width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
            height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
          },
        },
      },
    ];
  }
  const object = href === undefined ? undefined : context.embeddedObjects.get(href);
  if (!object) throw malformedEmbeddedObject("embedded object subdocument is missing", href);
  const shapeId = context.nextEmbeddedShapeId++;
  const anchorType = frameAnchorType(attributeString(frame, "text:anchor-type"));
  const styleName = attributeString(frame, "draw:style-name");
  object.frames = [
    ...(object.frames ?? []),
    {
      shapeId,
      ...(anchorType ? { anchorType } : {}),
      ...(styleName ? { styleName } : {}),
    },
  ];
  return [{ wpsShape: embeddedObjectShape(frame, context, shapeId) }];
}

function frameAnchorType(value: string | undefined): OdtFrameAnchorType | undefined {
  const anchorTypes: OdtFrameAnchorType[] = ["as-char", "char", "page", "paragraph", "frame"];
  if (value === undefined) return undefined;
  if (anchorTypes.includes(value as OdtFrameAnchorType)) return value as OdtFrameAnchorType;
  throw malformedEmbeddedObject(`unsupported frame anchor type: ${value}`);
}

function embeddedObjectShape(frame: Element, context: ParseContext, shapeId: number): ShapeOptions {
  const styleName = attributeString(frame, "draw:style-name");
  const graphic = styleName ? context.graphicStyles.get(styleName) : undefined;
  const fill = graphicFill(graphic);
  const outline = graphicOutline(graphic);
  const name = attributeString(frame, "draw:name");
  const x = lengthToEmu(attributeString(frame, "svg:x"));
  const y = lengthToEmu(attributeString(frame, "svg:y"));
  return {
    children: [],
    geometry: "rect",
    transformation: {
      ...(x || y ? { offset: { ...(x ? { x } : {}), ...(y ? { y } : {}) } } : {}),
      width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
      height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
    },
    nonVisualProperties: { id: shapeId, ...(name ? { name } : {}) },
    ...(fill ? { fill } : {}),
    ...(outline ? { outline } : {}),
  };
}

function malformedEmbeddedObject(
  reason: string,
  href?: string,
  structuredReason = reason,
): OdtParseError {
  return new OdtParseError(
    `content.xml: ${href ?? "draw:object"}: ${reason}`,
    "content.xml",
    "/draw:frame/draw:object/@xlink:href",
    "draw:object",
    structuredReason,
  );
}

/** draw:custom-shape maps back to the shared docx shape model. */
export function parseCustomShape(
  element: Element,
  context: ParseContext,
): { wpsShape: ShapeOptions } | undefined {
  const styleName = attributeString(element, "draw:style-name");
  const graphic = styleName ? context.graphicStyles.get(styleName) : undefined;
  const enhanced = childNamed(element, "draw:enhanced-geometry");
  const presetType = attributeString(enhanced, "draw:type");
  const preset = presetType ? PRESET_GEOMETRY_DOCX[presetType] : undefined;
  const x = lengthToEmu(attributeString(element, "svg:x")) ?? 0;
  const y = lengthToEmu(attributeString(element, "svg:y")) ?? 0;
  const width = lengthToEmu(attributeString(element, "svg:width")) ?? 0;
  const height = lengthToEmu(attributeString(element, "svg:height")) ?? 0;
  const fill: FillOptions | undefined = graphicFill(graphic);
  const outline: OutlineOptions | undefined = graphicOutline(graphic);
  const name = attributeString(element, "draw:name");
  return {
    wpsShape: {
      children: [],
      transformation: {
        ...(x || y ? { offset: { ...(x ? { x } : {}), ...(y ? { y } : {}) } } : {}),
        width,
        height,
      },
      ...(name ? { altText: { name } } : {}),
      ...(preset ? { geometry: preset } : {}),
      ...(fill ? { fill } : {}),
      ...(outline ? { outline } : {}),
    },
  };
}

export function parseTextBoxFrame(frame: Element, context: ParseContext): ParagraphChild {
  const textBox = childNamed(frame, "draw:text-box");
  return {
    wpsShape: {
      ...parseShapeElement(frame, context),
      children: childrenNamed(textBox, "text:p").map((paragraph) => ({
        paragraph: { text: textOf(paragraph) },
      })),
      bodyPropertiesPresent: true,
    },
  };
}

export function parseGroup(element: Element, context: ParseContext): ParagraphChild {
  const childElements = (element.elements ?? []).filter(
    (child) => isShapeElement(child.name) || child.name === "draw:g",
  );
  const children: WpgChildMediaData[] = childElements.map((child) => {
    if (child.name === "draw:g") {
      const group = parseGroup(child, context);
      if (!("wpgGroup" in group)) throw unknownOdtElement(child);
      return {
        type: "wpg",
        transformation: group.wpgGroup.transformation,
        children: group.wpgGroup.children,
      };
    }
    const shape = parseShapeElement(child, context);
    return {
      type: "wps",
      transformation: shape.transformation,
      data: shape,
      ...(shape.fill ? { fill: shape.fill } : {}),
      ...(shape.outline ? { outline: shape.outline } : {}),
    };
  });
  const transformations = [
    ...children.map((child) => child.transformation),
    ...childElements
      .filter((child) => child.name === "draw:g")
      .map((child) => {
        const group = parseGroup(child, context);
        return "wpgGroup" in group ? group.wpgGroup.transformation : { width: 0, height: 0 };
      }),
  ];
  const left = Math.min(
    ...transformations.map((transformation) => Number(transformation.offset?.x ?? 0)),
  );
  const top = Math.min(
    ...transformations.map((transformation) => Number(transformation.offset?.y ?? 0)),
  );
  const right = Math.max(
    ...transformations.map(
      (transformation) => Number(transformation.offset?.x ?? 0) + Number(transformation.width),
    ),
  );
  const bottom = Math.max(
    ...transformations.map(
      (transformation) => Number(transformation.offset?.y ?? 0) + Number(transformation.height),
    ),
  );
  return {
    wpgGroup: {
      children,
      transformation: {
        offset: { x: left, y: top },
        width: right - left,
        height: bottom - top,
      },
      childOffsetX: left,
      childOffsetY: top,
      childExtentWidth: right - left,
      childExtentHeight: bottom - top,
      altText: attributeString(element, "draw:name")
        ? { name: attributeString(element, "draw:name") ?? "" }
        : undefined,
    } satisfies GroupOptions,
  };
}
