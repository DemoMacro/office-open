import {
  toUint8Array,
  type ChartSpaceOptions,
  type EndpointConnectionOptions,
} from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  lengthToEmu,
  xmlElement,
} from "@office-open/odf";
import type { ChartOptions, PictureOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";
import { toOdfLength } from "./shape";

/** Binary image collected during generation — emitted as a Pictures/ entry. */
export interface OdpImage {
  path: string;
  data: Uint8Array;
}

/** Embedded chart subdocument collected during generation. */
export interface OdpChart {
  path: string;
  chart: ChartOptions;
}

/** Slide picture renders as a positioned draw:frame + draw:image. */
export function pictureFrameXml(picture: PictureOptions, images: OdpImage[]): string {
  let path: string;
  if (picture.data === undefined) {
    if (!picture.sourceUrl) return "";
    path = picture.sourceUrl;
  } else {
    const data = toUint8Array(picture.data);
    path = `Pictures/image${images.length + 1}.${picture.type}`;
    images.push({ path, data });
  }
  return xmlElement(
    "draw:frame",
    {
      "draw:name": picture.name,
      "svg:x": toOdfLength(picture.x),
      "svg:y": toOdfLength(picture.y),
      "svg:width": toOdfLength(picture.width),
      "svg:height": toOdfLength(picture.height),
    },
    [
      xmlElement("draw:image", {
        "xlink:href": path,
        ...(picture.sourceUrl && picture.data === undefined
          ? { "xlink:type": "simple", "xlink:show": "embed" }
          : {}),
      }),
    ],
  );
}

/** Slide chart renders as a positioned draw:frame + embedded chart object. */
export function chartFrameXml(chart: ChartOptions, charts: OdpChart[]): string {
  const path = `Object ${charts.length + 1}`;
  charts.push({ path, chart });
  return xmlElement(
    "draw:frame",
    {
      "draw:name": chart.name,
      "svg:x": toOdfLength(chart.x),
      "svg:y": toOdfLength(chart.y),
      "svg:width": toOdfLength(chart.width),
      "svg:height": toOdfLength(chart.height),
    },
    [
      xmlElement("draw:object", {
        "xlink:href": `./${path}`,
        "xlink:type": "simple",
        "xlink:show": "embed",
        "xlink:actuate": "onLoad",
      }),
    ],
  );
}

/** Parses a draw:frame + embedded chart subdocument back to a slide chart. */
export function parseChartFrame(
  frame: Element,
  chartPool: Map<string, ChartSpaceOptions>,
): { chart: ChartOptions } | undefined {
  const object = childNamed(frame, "draw:object");
  if (!object) return undefined;
  const href = attributeString(object, "xlink:href")?.replace(/^\.\//, "").replace(/\/$/, "");
  const name = attributeString(frame, "draw:name") ?? href;
  const path = `/draw:frame[@draw:name="${name ?? ""}"]/draw:object`;
  if (!href || !name) {
    const reason = "chart frame has no object reference";
    throw new OdpParseError(
      `content.xml: ${path}: draw:object: ${reason}`,
      "content.xml",
      path,
      "draw:object",
      reason,
    );
  }
  const chart = chartPool.get(href);
  if (!chart) {
    const reason = "referenced chart subdocument is missing";
    throw new OdpParseError(
      `content.xml: ${path}: ${href}: ${reason}`,
      "content.xml",
      path,
      href,
      reason,
    );
  }
  return {
    chart: {
      ...chart,
      name,
      x: lengthToEmu(attributeString(frame, "svg:x")),
      y: lengthToEmu(attributeString(frame, "svg:y")),
      width: lengthToEmu(attributeString(frame, "svg:width")),
      height: lengthToEmu(attributeString(frame, "svg:height")),
    },
  };
}

/** Reads a draw:{start,end}-shape/-glue-point pair as a glued endpoint. */
export function endpointConnection(
  element: Element,
  end: "start" | "end",
): EndpointConnectionOptions | undefined {
  const id = Number(attributeString(element, `draw:${end}-shape`));
  const index = attributeNumber(element, `draw:${end}-glue-point`);
  return Number.isInteger(id) && id > 0 && index !== undefined ? { id, index } : undefined;
}

/** Parses a draw:frame + draw:image back to a slide picture child. */
export function parsePictureFrame(
  frame: Element,
  binaries: Record<string, Uint8Array>,
  usedBinaryPaths: Set<string>,
): { picture: PictureOptions } | undefined {
  const image = childNamed(frame, "draw:image");
  if (!image) return undefined;
  const path = attributeString(image, "xlink:href")?.replace(/^\//, "");
  const data = path ? binaries[path] : undefined;
  if (data && path) usedBinaryPaths.add(path);
  if (!data || !path) {
    const href = attributeString(image, "xlink:href");
    if (!href) return undefined;
    if (isExternalUrl(href))
      return {
        picture: {
          type: pictureTypeFromUrl(href),
          sourceUrl: href,
          x: lengthToEmu(attributeString(frame, "svg:x")),
          y: lengthToEmu(attributeString(frame, "svg:y")),
          width: lengthToEmu(attributeString(frame, "svg:width")),
          height: lengthToEmu(attributeString(frame, "svg:height")),
        },
      };
    const reason = "missing-package-part";
    const referencePath = "/draw:frame/draw:image/@xlink:href";
    throw new OdpParseError(
      `content.xml: ${href}: package image is missing`,
      "content.xml",
      referencePath,
      "draw:image",
      reason,
    );
  }
  return {
    picture: {
      type: (path.split(".").pop() ?? "png") as PictureOptions["type"],
      data,
      x: lengthToEmu(attributeString(frame, "svg:x")),
      y: lengthToEmu(attributeString(frame, "svg:y")),
      width: lengthToEmu(attributeString(frame, "svg:width")),
      height: lengthToEmu(attributeString(frame, "svg:height")),
    },
  };
}

function isExternalUrl(value: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("../");
}

function pictureTypeFromUrl(value: string): PictureOptions["type"] {
  const extension = value.split(/[?#]/, 1)[0]?.split("/").pop()?.split(".").pop()?.toLowerCase();
  return extension === "jpeg"
    ? "jpg"
    : ((["png", "jpg", "gif", "bmp", "emf", "wmf"].includes(extension ?? "")
        ? extension
        : "png") as PictureOptions["type"]);
}
