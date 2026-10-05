import { toUint8Array, type ChartSpaceOptions, type FormContainerOptions } from "@office-open/core";
import type { FillOptions, OutlineOptions } from "@office-open/core/drawing";
import type {
  ChartOptions,
  ParagraphChild,
  PictureOptions,
  RunOptions,
  SectionChild,
  ShapeOptions,
} from "@office-open/docx";
import {
  attributeString,
  CHART_MIME,
  childNamed,
  childrenNamed,
  emuToLength,
  graphicFill,
  graphicOutline,
  lengthToEmu,
  PRESET_GEOMETRY_DOCX,
  presetGeometryOdf,
  pushShapeStyle,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { unknownOdtElement, type ParseContext } from "./body";
import { OdtParseError } from "./error";

export type PictureType = NonNullable<PictureOptions["type"]>;

export type RasterPictureType = Exclude<PictureType, "svg">;

export type CanonicalFormControl = NonNullable<FormContainerOptions["controls"][number]>;

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

/** Inline chart renders as a draw:frame + draw:object pointing at the subdocument. */
export function chartFrameXml(chart: ChartOptions, charts: OdtChart[]): string {
  const path = `Object ${charts.length + 1}`;
  charts.push({ path, chart });
  return xmlElement(
    "draw:frame",
    {
      "text:anchor-type": "as-char",
      "svg:width": emuToLength(chart.transformation.width),
      "svg:height": emuToLength(chart.transformation.height),
    },
    [xmlElement("draw:object", { "xlink:href": `./${path}`, "xlink:type": "simple" })],
  );
}

/** Inline shape renders as a positioned draw:custom-shape with preset geometry. */
export function wpsShapeFrameXml(shape: ShapeOptions, styles: string[]): string {
  const geometry = typeof shape.geometry === "string" ? { preset: shape.geometry } : shape.geometry;
  const offset = shape.transformation.offset;
  return xmlElement(
    "draw:custom-shape",
    {
      "text:anchor-type": "as-char",
      "draw:style-name": addShapeStyle(shape, styles),
      "svg:x":
        offset?.left !== undefined
          ? typeof offset.left === "number"
            ? emuToLength(offset.left)
            : offset.left
          : undefined,
      "svg:y":
        offset?.top !== undefined
          ? typeof offset.top === "number"
            ? emuToLength(offset.top)
            : offset.top
          : undefined,
      "svg:width":
        typeof shape.transformation.width === "number"
          ? emuToLength(shape.transformation.width)
          : shape.transformation.width,
      "svg:height":
        typeof shape.transformation.height === "number"
          ? emuToLength(shape.transformation.height)
          : shape.transformation.height,
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
      "svg:width": emuToLength(picture.transformation.width),
      "svg:height": emuToLength(picture.transformation.height),
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
  const name = attributeString(element, "draw:name");
  return {
    children: [],
    geometry: element.name === "draw:rect" ? "rect" : "ellipse",
    ...(name ? { altText: { name } } : {}),
    transformation: {
      ...(lengthToEmu(attributeString(element, "svg:x"))
        ? { offset: { left: lengthToEmu(attributeString(element, "svg:x")) } }
        : {}),
      ...(lengthToEmu(attributeString(element, "svg:y"))
        ? { offset: { top: lengthToEmu(attributeString(element, "svg:y")) } }
        : {}),
      width: lengthToEmu(attributeString(element, "svg:width")) ?? 0,
      height: lengthToEmu(attributeString(element, "svg:height")) ?? 0,
    },
  };
}

export function parseShapeHyperlink(element: Element): ParagraphChild {
  return {
    hyperlink: {
      url: attributeString(element, "xlink:href") ?? "",
      children: childrenNamed(element, "draw:ellipse")
        .concat(childrenNamed(element, "draw:rect"))
        .map((shape) => ({ wpsShape: parseShape(shape) })),
    },
  };
}

/** draw:frame + draw:image maps back to an inline picture run. */
export function parsePictureFrame(frame: Element, context: ParseContext): ParagraphChild[] {
  const href = attributeString(childNamed(frame, "draw:image"), "xlink:href");
  const path = href?.replace(/^\//, "");
  const data = path ? context.binaries[path] : undefined;
  if (!data || !path) {
    if (href && isExternalUrl(href))
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
    if (href)
      throw new OdtParseError(
        `content.xml: ${href}: package image is missing`,
        "content.xml",
        "/draw:frame/draw:image/@xlink:href",
        "draw:image",
        "missing-package-part",
      );
    return [];
  }
  const extension = path.split(".").pop() ?? "png";
  const imageType = extension === "jpeg" ? "jpg" : extension;
  if (!isRasterPictureType(imageType)) {
    throw unknownOdtElement(frame);
  }
  return [
    {
      picture: {
        type: imageType,
        data,
        transformation: {
          width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
          height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
        },
      },
    },
  ];
}

export function isExternalUrl(value: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("../");
}

function pictureTypeFromUrl(value: string): RasterPictureType {
  const extension = value.split(/[?#]/, 1)[0]?.split("/").pop()?.split(".").pop()?.toLowerCase();
  if (extension === "jpeg") return "jpg";
  return isRasterPictureType(extension ?? "") ? (extension as RasterPictureType) : "png";
}

/** draw:frame + draw:object resolves an embedded chart subdocument. */
export function parseChartFrame(frame: Element, context: ParseContext): RunOptions[] {
  const href = attributeString(childNamed(frame, "draw:object"), "xlink:href")
    ?.replace(/^\.\//, "")
    .replace(/^\//, "");
  const chart = href ? context.chartBodies.get(href) : undefined;
  if (!chart) {
    const mediaType = href ? context.objectMediaTypes.get(href) : undefined;
    const reason = !href
      ? "chart frame has no object reference"
      : mediaType === undefined
        ? "embedded object media type is not declared"
        : mediaType === CHART_MIME
          ? "referenced chart subdocument is missing"
          : "embedded object has no canonical ODT mapping";
    throw new OdtParseError(
      `content.xml: ${href ?? "draw:object"}: ${reason}`,
      "content.xml",
      "/draw:frame/draw:object/@xlink:href",
      "draw:object",
      reason,
    );
  }
  return [
    {
      chart: {
        ...chart,
        transformation: {
          width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
          height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
        },
      },
    } as RunOptions,
  ];
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
        ...(x || y ? { offset: { ...(x ? { left: x } : {}), ...(y ? { top: y } : {}) } } : {}),
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
