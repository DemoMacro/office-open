import {
  toUint8Array,
  type ChartSpaceOptions,
  type EndpointConnectionOptions,
} from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  isExternalOdfUrl,
  lengthToEmu,
  resolveEmbeddedChart,
  xmlElement,
} from "@office-open/odf";
import type {
  AudioType,
  AudioFrameOptions,
  ChartOptions,
  MediaPlaybackOptions,
  PictureOptions,
  VideoType,
  VideoFrameOptions,
} from "@office-open/pptx";
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

/** Slide media renders as a positioned draw:frame + draw:plugin. */
export function mediaFrameXml(
  media: VideoFrameOptions | AudioFrameOptions,
  mediaFiles: OdpImage[],
  images: OdpImage[],
): string {
  const extension = media.type ?? "mp4";
  const path =
    media.data === undefined
      ? (media.sourceUrl ?? "")
      : `Media/media${mediaFiles.length + 1}.${extension}`;
  if (media.data !== undefined) mediaFiles.push({ path, data: toUint8Array(media.data) });
  let posterPath: string | undefined;
  if (media.poster !== undefined) {
    const poster = toUint8Array(media.poster);
    posterPath = `Pictures/poster${images.length + 1}.${media.posterType ?? "png"}`;
    images.push({ path: posterPath, data: poster });
  }
  const playback = media.playback;
  const parameter = (name: string, value: string) =>
    xmlElement("draw:param", { "draw:name": name, "draw:value": value });
  return xmlElement(
    "draw:frame",
    {
      "draw:name": media.name,
      "svg:x": toOdfLength(media.x),
      "svg:y": toOdfLength(media.y),
      "svg:width": toOdfLength(media.width),
      "svg:height": toOdfLength(media.height),
    },
    [
      xmlElement(
        "draw:plugin",
        {
          "xlink:href": path,
          "xlink:type": "simple",
          "xlink:show": "embed",
          "xlink:actuate": "onLoad",
          "draw:mime-type": "application/vnd.sun.star.media",
        },
        [
          ...(playback?.loop === undefined ? [] : [parameter("Loop", String(playback.loop))]),
          ...(playback?.mute === undefined ? [] : [parameter("Mute", String(playback.mute))]),
          ...(playback?.volume === undefined
            ? []
            : [parameter("VolumeDB", String(playback.volume))]),
          ...(playback?.zoom === undefined ? [] : [parameter("Zoom", playback.zoom)]),
        ],
      ),
      ...(posterPath
        ? [xmlElement("draw:image", { "xlink:href": posterPath, "xlink:type": "simple" })]
        : []),
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
  const chart = href === undefined ? undefined : resolveEmbeddedChart(chartPool, href);
  if (!chart) return undefined;
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
    if (isExternalOdfUrl(href))
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

function pictureTypeFromUrl(value: string): PictureOptions["type"] {
  const extension = value.split(/[?#]/, 1)[0]?.split("/").pop()?.split(".").pop()?.toLowerCase();
  if (extension === "jpeg") return "jpg";
  const supported = ["png", "jpg", "gif", "bmp", "emf", "wmf"].includes(extension ?? "");
  return (supported ? extension : "png") as PictureOptions["type"];
}

/** Parses a draw:frame + media draw:plugin back to a slide video/audio child. */
export function parseMediaFrame(
  frame: Element,
  binaries: Record<string, Uint8Array>,
  usedBinaryPaths: Set<string>,
): { video: VideoFrameOptions } | { audio: AudioFrameOptions } | undefined {
  const plugin = childNamed(frame, "draw:plugin");
  const mimeType = attributeString(plugin, "draw:mime-type");
  if (!plugin || mimeType !== "application/vnd.sun.star.media") return undefined;
  const href = attributeString(plugin, "xlink:href")?.replace(/^\//, "");
  if (!href) return undefined;
  const extension = href.split(/[?#]/, 1)[0]?.split(".").pop()?.toLowerCase() ?? "";
  if (!isAudioType(extension) && !isVideoType(extension)) return undefined;
  const data = binaries[href];
  if (data) usedBinaryPaths.add(href);
  else if (!isExternalOdfUrl(href)) throw missingMediaPart(href);
  const name = attributeString(frame, "draw:name");
  const x = lengthToEmu(attributeString(frame, "svg:x"));
  const y = lengthToEmu(attributeString(frame, "svg:y"));
  const width = lengthToEmu(attributeString(frame, "svg:width"));
  const height = lengthToEmu(attributeString(frame, "svg:height"));
  const poster = childNamed(frame, "draw:image");
  const posterHref = poster ? attributeString(poster, "xlink:href")?.replace(/^\//, "") : undefined;
  const posterData = posterHref ? binaries[posterHref] : undefined;
  if (posterData && posterHref) usedBinaryPaths.add(posterHref);
  if (isAudioType(extension)) {
    const geometry: MediaGeometry = {};
    assignMediaGeometry(geometry, { name, x, y, width, height });
    const audio: AudioFrameOptions = { ...geometry, type: extension };
    if (data === undefined) audio.sourceUrl = href;
    else {
      audio.data = data;
      audio.fileName = href.split("/").pop();
    }
    audio.playback = parseMediaPlayback(plugin);
    return { audio };
  }
  const geometry: MediaGeometry = {};
  assignMediaGeometry(geometry, { name, x, y, width, height });
  const video: VideoFrameOptions = { ...geometry, type: extension };
  if (data === undefined) video.sourceUrl = href;
  else {
    video.data = data;
    video.fileName = href.split("/").pop();
  }
  video.playback = parseMediaPlayback(plugin);
  if (posterData && posterHref) {
    video.poster = posterData;
    video.posterType = posterType(posterHref);
  }
  return { video };
}

interface MediaGeometry {
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

function assignMediaGeometry(
  options: MediaGeometry,
  frame: {
    name: string | undefined;
    x: number | undefined;
    y: number | undefined;
    width: number | undefined;
    height: number | undefined;
  },
): void {
  if (frame.name !== undefined) options.name = frame.name;
  if (frame.x !== undefined) options.x = frame.x;
  if (frame.y !== undefined) options.y = frame.y;
  if (frame.width !== undefined) options.width = frame.width;
  if (frame.height !== undefined) options.height = frame.height;
}

function isAudioType(value: string): value is AudioType {
  return ["mp3", "wav", "wma", "aac"].includes(value);
}

function isVideoType(value: string): value is VideoType {
  return ["mp4", "mov", "wmv", "avi"].includes(value);
}

function posterType(path: string): VideoFrameOptions["posterType"] {
  const extension = path.split(".").pop() ?? "png";
  switch (extension) {
    case "jpg":
    case "gif":
    case "bmp":
    case "tif":
    case "ico":
    case "emf":
    case "wmf":
      return extension;
    default:
      return "png";
  }
}

function missingMediaPart(path: string): OdpParseError {
  const reason = "missing-package-part";
  return new OdpParseError(
    `content.xml: ${path}: package media is missing`,
    "content.xml",
    "/draw:frame/draw:plugin/@xlink:href",
    "draw:plugin",
    reason,
  );
}

function parseMediaPlayback(plugin: Element): MediaPlaybackOptions | undefined {
  const params = (plugin.elements ?? []).filter((child) => child.name === "draw:param");
  const value = (name: string) => {
    const param = params.find((child) => attributeString(child, "draw:name") === name);
    return param === undefined ? undefined : attributeString(param, "draw:value");
  };
  const loop = value("Loop");
  const mute = value("Mute");
  const volume = value("VolumeDB");
  const zoom = value("Zoom");
  const playback: MediaPlaybackOptions = {
    ...(loop === undefined ? {} : { loop: loop === "true" }),
    ...(mute === undefined ? {} : { mute: mute === "true" }),
    ...(volume === undefined || Number.isNaN(Number(volume)) ? {} : { volume: Number(volume) }),
    ...(zoom === undefined ? {} : { zoom }),
  };
  return Object.keys(playback).length > 0 ? playback : undefined;
}
