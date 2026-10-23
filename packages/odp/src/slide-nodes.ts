import type { TextBodyOptions } from "@office-open/core";
import type { FormEventListenerOptions } from "@office-open/odf";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  odfLengthToEmu,
  textOf,
  unknownAttributeNames,
  unknownChildElement,
  xmlElement,
} from "@office-open/odf";
import type { XmlAttributes } from "@office-open/odf";
import type { SlideChild, TableOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";
import { parseSlideTable, slideTableXml } from "./slide";
import { parseParagraph, textBodyXml } from "./text-run";

type Length = number | `${number}${"mm" | "cm" | "in" | "pt" | "pc" | "pi" | "px"}`;
type Distance = Length | `${number}%`;

/** ODF three-dimensional coordinate expressed as structured lengths. */
export interface OdpVector3DOptions {
  x: Length;
  y: Length;
  z: Length;
}

/** Shared drawing identity, geometry, style, anchoring, and presentation fields. */
export interface OdpSlideNodeBaseOptions {
  xmlId?: string;
  drawId?: string;
  name?: string;
  zIndex?: number;
  layer?: string;
  drawStyleName?: string;
  drawClassNames?: string;
  presentationStyleName?: string;
  presentationClassNames?: string;
  transform?: string;
  captionId?: string;
  textStyleName?: string;
  tableEndCellAddress?: string;
  tableEndX?: Length;
  tableEndY?: Length;
  tableBackground?: boolean;
  anchorType?: "page" | "frame" | "paragraph" | "char" | "as-char";
  anchorPageNumber?: number;
  presentationClass?: string;
  presentationPlaceholder?: boolean;
  presentationUserTransformed?: boolean;
}

export interface OdpGluePointOptions {
  id: number;
  x: Distance;
  y: Distance;
  align?:
    | "top-left"
    | "top"
    | "top-right"
    | "left"
    | "center"
    | "right"
    | "bottom-left"
    | "bottom-right";
  escapeDirection: "auto" | "left" | "right" | "top" | "bottom";
}

export interface OdpLightOptions {
  direction: OdpVector3DOptions;
  diffuseColor?: string;
  enabled?: boolean;
  specular?: boolean;
}

export interface OdpCubeOptions extends OdpSlideNodeBaseOptions {
  type: "cube";
  minEdge?: OdpVector3DOptions;
  maxEdge?: OdpVector3DOptions;
}

export interface OdpSphereOptions extends OdpSlideNodeBaseOptions {
  type: "sphere";
  center?: OdpVector3DOptions;
  size?: OdpVector3DOptions;
}

export interface OdpExtrudeOptions extends OdpSlideNodeBaseOptions {
  type: "extrude";
  path: string;
  viewBox: [number, number, number, number];
  depth?: Length;
  closeFront?: boolean;
  closeBack?: boolean;
}

export interface OdpRotateOptions extends OdpSlideNodeBaseOptions {
  type: "rotate";
  path: string;
  viewBox: [number, number, number, number];
  horizontal?: boolean;
}

export type OdpSceneObjectOptions =
  | OdpCubeOptions
  | OdpSphereOptions
  | OdpExtrudeOptions
  | OdpRotateOptions
  | OdpSceneOptions;

/** Complete dr3d:scene projection; scene objects may be nested one level. */
export interface OdpSceneOptions extends OdpSlideNodeBaseOptions {
  x?: Length;
  y?: Length;
  width?: Length;
  height?: Length;
  viewRecommendationPoint?: OdpVector3DOptions;
  viewPlaneNormal?: OdpVector3DOptions;
  viewUp?: OdpVector3DOptions;
  projection?: "parallel" | "perspective";
  distance?: Length;
  focalLength?: Length;
  shadowSlant?: number;
  shadeMode?: "flat" | "phong" | "gouraud" | "draft";
  ambientColor?: string;
  lightingMode?: boolean;
  title?: string;
  description?: string;
  lights?: OdpLightOptions[];
  objects?: OdpSceneObjectOptions[];
  gluePoints?: OdpGluePointOptions[];
}

/** Complete draw:measure projection. */
export interface OdpMeasureOptions extends OdpSlideNodeBaseOptions {
  x1: Length;
  y1: Length;
  x2: Length;
  y2: Length;
  title?: string;
  description?: string;
  listeners?: FormEventListenerOptions[];
  gluePoints?: OdpGluePointOptions[];
  measureAlign?: "left" | "right" | "automatic";
  measureVerticalAlign?: "automatic" | "above" | "below";
  text?: TextBodyOptions;
}

/** Complete draw:page-thumbnail projection. */
export interface OdpPageThumbnailOptions extends OdpSlideNodeBaseOptions {
  x?: Length;
  y?: Length;
  width?: Length;
  height?: Length;
  pageNumber?: number;
  title?: string;
  description?: string;
}

export interface OdpDataReferenceOptions {
  href: string;
  show?: "embed";
  actuate?: "onLoad";
}

export interface OdpParameterOptions {
  name?: string;
  value?: string;
}

export interface OdpAppletOptions {
  xmlId?: string;
  code?: string;
  object?: string;
  archive?: string;
  mayScript?: boolean;
  data?: OdpDataReferenceOptions;
  parameters?: OdpParameterOptions[];
}

export interface OdpFloatingFrameOptions {
  xmlId?: string;
  frameName?: string;
  data: OdpDataReferenceOptions;
}

export interface OdpPluginOptions {
  xmlId?: string;
  mimeType?: string;
  data: OdpDataReferenceOptions;
  parameters?: OdpParameterOptions[];
}

export interface OdpTextBoxOptions {
  chainNextName?: string;
  cornerRadius?: Length;
  minHeight?: Length;
  maxHeight?: Length;
  minWidth?: Length;
  maxWidth?: Length;
  autoGrowHeight?: boolean;
  autoGrowWidth?: boolean;
  text: TextBodyOptions;
}

export type OdpFrameContentOptions =
  | { kind: "applet"; applet: OdpAppletOptions }
  | { kind: "floating-frame"; floatingFrame: OdpFloatingFrameOptions }
  | { kind: "plugin"; plugin: OdpPluginOptions }
  | { kind: "table"; table: TableOptions }
  | { kind: "text-box"; textBox: OdpTextBoxOptions };

/** Typed fallback for an RNG-valid draw:frame not projected to a PPTX frame. */
export interface OdpFrameOptions extends OdpSlideNodeBaseOptions {
  x?: Length;
  y?: Length;
  width?: Length;
  height?: Length;
  relativeWidth?: string;
  relativeHeight?: string;
  copyOf?: string;
  title?: string;
  description?: string;
  listeners?: FormEventListenerOptions[];
  gluePoints?: OdpGluePointOptions[];
  contents: OdpFrameContentOptions[];
}

export type OdpSlideChild =
  | SlideChild
  | { frame: OdpFrameOptions }
  | { measure: OdpMeasureOptions }
  | { pageThumbnail: OdpPageThumbnailOptions }
  | { scene: OdpSceneOptions };

export type OdpSlideNodePayload = Extract<
  OdpSlideChild,
  | { frame: OdpFrameOptions }
  | { measure: OdpMeasureOptions }
  | { pageThumbnail: OdpPageThumbnailOptions }
  | { scene: OdpSceneOptions }
>;

/** Typed ODF-only child preserved at its visible source position. */
export interface OdpSlideNodeOptions {
  position: number;
  node: OdpSlideNodePayload;
}

const APPLET_ATTRIBUTES = ["draw:archive", "draw:code", "draw:may-script", "draw:object", "xml:id"];
const DATA_ATTRIBUTES = ["xlink:actuate", "xlink:href", "xlink:show", "xlink:type"];
const FLOATING_FRAME_ATTRIBUTES = ["draw:frame-name", "xml:id"];
const GLUE_POINT_ATTRIBUTES = ["draw:align", "draw:escape-direction", "draw:id", "svg:x", "svg:y"];
const PARAMETER_ATTRIBUTES = ["draw:name", "draw:value"];
const PLUGIN_ATTRIBUTES = ["draw:mime-type", "xml:id"];
const SCENE_OBJECT_ATTRIBUTES = [
  "draw:id",
  "draw:layer",
  "draw:style-name",
  "draw:transform",
  "draw:z-index",
  "dr3d:transform",
  "xml:id",
];

/** Parse the raw ODF nodes that have no PPTX canonical equivalent. */
export function parseFrameNode(element: Element): { frame: OdpFrameOptions } {
  const path = "/draw:frame";
  assertKnown(
    element,
    [...shapeAttributeNames(true), "draw:copy-of", "style:rel-height", "style:rel-width"],
    path,
  );
  assertKnownChildren(
    element,
    [
      "draw:applet",
      "draw:floating-frame",
      "draw:glue-point",
      "draw:plugin",
      "draw:text-box",
      "office:event-listeners",
      "svg:desc",
      "svg:title",
      "table:table",
    ],
    path,
  );
  return {
    frame: {
      ...parseBase(element),
      ...parseGeometry(element),
      ...(attributeString(element, "style:rel-width") === undefined
        ? {}
        : { relativeWidth: attributeString(element, "style:rel-width") }),
      ...(attributeString(element, "style:rel-height") === undefined
        ? {}
        : { relativeHeight: attributeString(element, "style:rel-height") }),
      ...(attributeString(element, "draw:copy-of") === undefined
        ? {}
        : { copyOf: attributeString(element, "draw:copy-of") }),
      ...parseTitleDescription(element),
      ...(childNamed(element, "office:event-listeners") === undefined
        ? {}
        : { listeners: parseListeners(childNamed(element, "office:event-listeners")!) }),
      ...parseGluePoints(element),
      contents: (element.elements ?? [])
        .filter((child) => child.type === "element")
        .flatMap((child): OdpFrameContentOptions[] => {
          if (child.name === "draw:applet")
            return [{ kind: "applet" as const, applet: parseApplet(child) }];
          if (child.name === "draw:floating-frame")
            return [{ kind: "floating-frame" as const, floatingFrame: parseFloatingFrame(child) }];
          if (child.name === "draw:plugin")
            return [{ kind: "plugin" as const, plugin: parsePlugin(child) }];
          if (child.name === "draw:text-box")
            return [{ kind: "text-box" as const, textBox: parseTextBox(child) }];
          if (child.name === "table:table")
            return [{ kind: "table" as const, table: parseTable(child) }];
          return [];
        }),
    },
  };
}

export function parseMeasureNode(element: Element): { measure: OdpMeasureOptions } {
  const path = "/draw:measure";
  assertKnown(element, shapeAttributeNames(true), path);
  assertKnownChildren(
    element,
    ["draw:glue-point", "office:event-listeners", "svg:desc", "svg:title", "text:p"],
    path,
  );
  const x1 = requiredLength(element, "svg:x1", path);
  const y1 = requiredLength(element, "svg:y1", path);
  const x2 = requiredLength(element, "svg:x2", path);
  const y2 = requiredLength(element, "svg:y2", path);
  return {
    measure: {
      ...parseBase(element),
      x1,
      y1,
      x2,
      y2,
      ...parseTitleDescription(element),
      ...(childNamed(element, "office:event-listeners") === undefined
        ? {}
        : { listeners: parseListeners(childNamed(element, "office:event-listeners")!) }),
      ...parseGluePoints(element),
      ...(childrenNamed(element, "text:p").length === 0
        ? {}
        : {
            text: {
              paragraphs: childrenNamed(element, "text:p").map((paragraph) =>
                parseParagraph(paragraph, new Map()),
              ),
            },
          }),
    },
  };
}

export function parsePageThumbnailNode(element: Element): {
  pageThumbnail: OdpPageThumbnailOptions;
} {
  const path = "/draw:page-thumbnail";
  assertKnown(element, shapeAttributeNames(true), path);
  assertKnownChildren(element, ["svg:desc", "svg:title"], path);
  return {
    pageThumbnail: {
      ...parseBase(element),
      ...parseGeometry(element),
      ...(attributeNumber(element, "draw:page-number") === undefined
        ? {}
        : { pageNumber: attributeNumber(element, "draw:page-number") }),
      ...parseTitleDescription(element),
    },
  };
}

export function parseSceneNode(element: Element): { scene: OdpSceneOptions } {
  const path = "/dr3d:scene";
  assertKnown(element, sceneAttributeNames(), path);
  assertKnownChildren(
    element,
    [
      "dr3d:cube",
      "dr3d:extrude",
      "dr3d:light",
      "dr3d:rotate",
      "dr3d:scene",
      "dr3d:sphere",
      "draw:glue-point",
      "svg:desc",
      "svg:title",
    ],
    path,
  );
  const lights = childrenNamed(element, "dr3d:light");
  const objects = (element.elements ?? [])
    .filter((child) => child.type === "element" && child.name?.startsWith("dr3d:"))
    .filter((child) => child.name !== "dr3d:light")
    .map(parseSceneObject);
  return {
    scene: {
      ...parseBase(element),
      ...parseGeometry(element),
      ...(attributeString(element, "dr3d:vrp") === undefined
        ? {}
        : { viewRecommendationPoint: parseVector(attributeString(element, "dr3d:vrp")!, path) }),
      ...(attributeString(element, "dr3d:vpn") === undefined
        ? {}
        : { viewPlaneNormal: parseVector(attributeString(element, "dr3d:vpn")!, path) }),
      ...(attributeString(element, "dr3d:vup") === undefined
        ? {}
        : { viewUp: parseVector(attributeString(element, "dr3d:vup")!, path) }),
      ...(attributeString(element, "dr3d:projection") === undefined
        ? {}
        : {
            projection: parseEnum(element, "dr3d:projection", ["parallel", "perspective"], path) as
              | "parallel"
              | "perspective",
          }),
      ...(attributeString(element, "dr3d:distance") === undefined
        ? {}
        : { distance: parseLength(attributeString(element, "dr3d:distance")!, path) }),
      ...(attributeString(element, "dr3d:focal-length") === undefined
        ? {}
        : { focalLength: parseLength(attributeString(element, "dr3d:focal-length")!, path) }),
      ...(attributeNumber(element, "dr3d:shadow-slant") === undefined
        ? {}
        : { shadowSlant: attributeNumber(element, "dr3d:shadow-slant") }),
      ...(attributeString(element, "dr3d:shade-mode") === undefined
        ? {}
        : {
            shadeMode: parseEnum(
              element,
              "dr3d:shade-mode",
              ["flat", "phong", "gouraud", "draft"],
              path,
            ) as OdpSceneOptions["shadeMode"],
          }),
      ...(attributeString(element, "dr3d:ambient-color") === undefined
        ? {}
        : { ambientColor: attributeString(element, "dr3d:ambient-color") }),
      ...(attributeString(element, "dr3d:lighting-mode") === undefined
        ? {}
        : { lightingMode: parseBoolean(element, "dr3d:lighting-mode", path) }),
      ...parseTitleDescription(element),
      ...(lights.length > 0 ? { lights: lights.map((light) => parseLight(light, path)) } : {}),
      ...(objects.length > 0 ? { objects } : {}),
      ...parseGluePoints(element),
    },
  };
}

/** Serialize one previously parsed ODP-only slide child. */
export function slideNodeXml(child: OdpSlideNodePayload, styles: string[]): string {
  if ("frame" in child) return frameXml(child.frame, styles);
  if ("measure" in child) return measureXml(child.measure, styles);
  if ("pageThumbnail" in child) return pageThumbnailXml(child.pageThumbnail);
  if ("scene" in child) return sceneXml(child.scene);
  return "";
}

/** Parse any ODF-only slide node into its typed payload. */
export function parseSlideNode(element: Element): OdpSlideNodePayload {
  if (element.name === "draw:frame") return parseFrameNode(element);
  if (element.name === "draw:measure") return parseMeasureNode(element);
  if (element.name === "draw:page-thumbnail") return parsePageThumbnailNode(element);
  if (element.name === "dr3d:scene") return parseSceneNode(element);
  throw new OdpParseError(
    `Unexpected ODP slide node: ${element.name}`,
    "content.xml",
    `/${element.name}`,
    element.name ?? "",
    "unsupported-slide-node",
  );
}

function frameXml(options: OdpFrameOptions, styles: string[]): string {
  const contents = options.contents.map((content) => {
    if (content.kind === "applet") return appletXml(content.applet);
    if (content.kind === "floating-frame") return floatingFrameXml(content.floatingFrame);
    if (content.kind === "plugin") return pluginXml(content.plugin);
    if (content.kind === "table") return slideTableXml(content.table, styles);
    return textBoxXml(content.textBox, styles);
  });
  return xmlElement(
    "draw:frame",
    {
      ...baseAttributes(options, "draw"),
      "svg:x": odfLength(options.x),
      "svg:y": odfLength(options.y),
      "svg:width": odfLength(options.width),
      "svg:height": odfLength(options.height),
      "style:rel-width": options.relativeWidth,
      "style:rel-height": options.relativeHeight,
      "draw:copy-of": options.copyOf,
    },
    [
      ...contents,
      ...(options.listeners === undefined ? [] : [listenersXml(options.listeners)]),
      ...(options.gluePoints ?? []).map(gluePointXml),
      ...(options.title === undefined ? [] : [xmlElement("svg:title", undefined, [options.title])]),
      ...(options.description === undefined
        ? []
        : [xmlElement("svg:desc", undefined, [options.description])]),
    ],
  );
}

function measureXml(options: OdpMeasureOptions, styles: string[]): string {
  return xmlElement(
    "draw:measure",
    {
      ...baseAttributes(options, "draw"),
      "svg:x1": odfLength(options.x1),
      "svg:y1": odfLength(options.y1),
      "svg:x2": odfLength(options.x2),
      "svg:y2": odfLength(options.y2),
    },
    [
      ...(options.listeners === undefined ? [] : [listenersXml(options.listeners)]),
      ...(options.gluePoints ?? []).map(gluePointXml),
      ...(options.title === undefined ? [] : [xmlElement("svg:title", undefined, [options.title])]),
      ...(options.description === undefined
        ? []
        : [xmlElement("svg:desc", undefined, [options.description])]),
      ...textBodyXml(options.text, styles),
    ],
  );
}

function pageThumbnailXml(options: OdpPageThumbnailOptions): string {
  return xmlElement(
    "draw:page-thumbnail",
    {
      ...baseAttributes(options, "draw"),
      "svg:x": odfLength(options.x),
      "svg:y": odfLength(options.y),
      "svg:width": odfLength(options.width),
      "svg:height": odfLength(options.height),
      "draw:page-number": options.pageNumber,
    },
    [
      ...(options.title === undefined ? [] : [xmlElement("svg:title", undefined, [options.title])]),
      ...(options.description === undefined
        ? []
        : [xmlElement("svg:desc", undefined, [options.description])]),
    ],
  );
}

function sceneXml(options: OdpSceneOptions): string {
  return xmlElement(
    "dr3d:scene",
    {
      ...baseAttributes(options, "draw"),
      "svg:x": odfLength(options.x),
      "svg:y": odfLength(options.y),
      "svg:width": odfLength(options.width),
      "svg:height": odfLength(options.height),
      "dr3d:vrp": vectorAttributes(options.viewRecommendationPoint),
      "dr3d:vpn": vectorAttributes(options.viewPlaneNormal),
      "dr3d:vup": vectorAttributes(options.viewUp),
      "dr3d:projection": options.projection,
      "dr3d:distance": odfLength(options.distance),
      "dr3d:focal-length": odfLength(options.focalLength),
      "dr3d:shadow-slant": options.shadowSlant,
      "dr3d:shade-mode": options.shadeMode,
      "dr3d:ambient-color": options.ambientColor,
      "dr3d:lighting-mode": options.lightingMode,
    },
    [
      ...(options.title === undefined ? [] : [xmlElement("svg:title", undefined, [options.title])]),
      ...(options.description === undefined
        ? []
        : [xmlElement("svg:desc", undefined, [options.description])]),
      ...(options.lights ?? []).map(lightXml),
      ...(options.objects ?? []).map(sceneObjectXml),
      ...(options.gluePoints ?? []).map(gluePointXml),
    ],
  );
}

function parseSceneObject(element: Element): OdpSceneObjectOptions {
  const path = `/dr3d:scene/${element.name}`;
  const common = SCENE_OBJECT_ATTRIBUTES;
  const base = parseBase(element);
  if (element.name === "dr3d:cube") {
    assertKnown(element, [...common, "dr3d:max-edge", "dr3d:min-edge"], path);
    return {
      ...base,
      type: "cube",
      ...(attributeString(element, "dr3d:min-edge") === undefined
        ? {}
        : { minEdge: parseVector(attributeString(element, "dr3d:min-edge")!, path) }),
      ...(attributeString(element, "dr3d:max-edge") === undefined
        ? {}
        : { maxEdge: parseVector(attributeString(element, "dr3d:max-edge")!, path) }),
    };
  }
  if (element.name === "dr3d:sphere") {
    assertKnown(element, [...common, "dr3d:center", "dr3d:size"], path);
    return {
      ...base,
      type: "sphere",
      ...(attributeString(element, "dr3d:center") === undefined
        ? {}
        : { center: parseVector(attributeString(element, "dr3d:center")!, path) }),
      ...(attributeString(element, "dr3d:size") === undefined
        ? {}
        : { size: parseVector(attributeString(element, "dr3d:size")!, path) }),
    };
  }
  if (element.name === "dr3d:extrude" || element.name === "dr3d:rotate") {
    const d = attributeString(element, "svg:d");
    if (!d) throw requiredAttributeError(path, "svg:d");
    return {
      ...base,
      type: element.name === "dr3d:extrude" ? "extrude" : "rotate",
      path: d,
      viewBox: parseViewBox(element, path),
      ...(element.name === "dr3d:extrude"
        ? {
            ...(attributeString(element, "dr3d:depth") === undefined
              ? {}
              : { depth: parseLength(attributeString(element, "dr3d:depth")!, path) }),
            ...(attributeString(element, "dr3d:close-front") === undefined
              ? {}
              : { closeFront: parseBoolean(element, "dr3d:close-front", path) }),
            ...(attributeString(element, "dr3d:close-back") === undefined
              ? {}
              : { closeBack: parseBoolean(element, "dr3d:close-back", path) }),
          }
        : attributeString(element, "dr3d:horizontal") === undefined
          ? {}
          : { horizontal: parseBoolean(element, "dr3d:horizontal", path) }),
    };
  }
  return parseSceneNode(element).scene;
}

function sceneObjectXml(options: OdpSceneObjectOptions): string {
  if (!("type" in options)) return sceneXml(options);
  if (options.type === "cube")
    return xmlElement("dr3d:cube", {
      ...baseAttributes(options, "dr3d"),
      "dr3d:min-edge": vectorAttributes(options.minEdge),
      "dr3d:max-edge": vectorAttributes(options.maxEdge),
    });
  if (options.type === "sphere")
    return xmlElement("dr3d:sphere", {
      ...baseAttributes(options, "dr3d"),
      "dr3d:center": vectorAttributes(options.center),
      "dr3d:size": vectorAttributes(options.size),
    });
  if (options.type === "extrude")
    return xmlElement("dr3d:extrude", {
      ...baseAttributes(options, "dr3d"),
      "svg:d": options.path,
      "svg:viewBox": options.viewBox.join(" "),
      "dr3d:depth": odfLength(options.depth),
      "dr3d:close-front": options.closeFront,
      "dr3d:close-back": options.closeBack,
    });
  return xmlElement("dr3d:rotate", {
    ...baseAttributes(options, "dr3d"),
    "svg:d": options.path,
    "svg:viewBox": options.viewBox.join(" "),
    "dr3d:horizontal": options.horizontal,
  });
}

function lightXml(options: OdpLightOptions): string {
  return xmlElement("dr3d:light", {
    "dr3d:diffuse-color": options.diffuseColor,
    "dr3d:direction": vectorAttributes(options.direction),
    "dr3d:enabled": options.enabled,
    "dr3d:specular": options.specular,
  });
}

function appletXml(options: OdpAppletOptions): string {
  return xmlElement(
    "draw:applet",
    {
      ...dataAttributes(options.data),
      "draw:code": options.code,
      "draw:object": options.object,
      "draw:archive": options.archive,
      "draw:may-script": options.mayScript,
      "xml:id": options.xmlId,
    },
    (options.parameters ?? []).map(parameterXml),
  );
}

function floatingFrameXml(options: OdpFloatingFrameOptions): string {
  return xmlElement("draw:floating-frame", {
    ...dataAttributes(options.data),
    "draw:frame-name": options.frameName,
    "xml:id": options.xmlId,
  });
}

function pluginXml(options: OdpPluginOptions): string {
  return xmlElement(
    "draw:plugin",
    {
      ...dataAttributes(options.data),
      "draw:mime-type": options.mimeType,
      "xml:id": options.xmlId,
    },
    (options.parameters ?? []).map(parameterXml),
  );
}

function textBoxXml(options: OdpTextBoxOptions, styles: string[]): string {
  return xmlElement(
    "draw:text-box",
    {
      "draw:chain-next-name": options.chainNextName,
      "draw:corner-radius": odfLength(options.cornerRadius),
      "fo:min-height": odfLength(options.minHeight),
      "fo:max-height": odfLength(options.maxHeight),
      "fo:min-width": odfLength(options.minWidth),
      "fo:max-width": odfLength(options.maxWidth),
      "draw:auto-grow-height": options.autoGrowHeight,
      "draw:auto-grow-width": options.autoGrowWidth,
    },
    textBodyXml(options.text, styles),
  );
}

function dataAttributes(options: OdpDataReferenceOptions | undefined): XmlAttributes {
  if (options === undefined) return {};
  return {
    "xlink:type": "simple",
    "xlink:href": options.href,
    "xlink:show": options.show,
    "xlink:actuate": options.actuate,
  };
}

function parameterXml(options: OdpParameterOptions): string {
  return xmlElement("draw:param", {
    "draw:name": options.name,
    "draw:value": options.value,
  });
}

function gluePointXml(options: OdpGluePointOptions): string {
  return xmlElement("draw:glue-point", {
    "draw:id": options.id,
    "svg:x": odfDistance(options.x),
    "svg:y": odfDistance(options.y),
    "draw:align": options.align,
    "draw:escape-direction": options.escapeDirection,
  });
}

function listenersXml(events: FormEventListenerOptions[]): string {
  return xmlElement(
    "office:event-listeners",
    undefined,
    events.map((event) =>
      event.target === "macro"
        ? xmlElement("script:event-listener", {
            "script:event-name": event.eventName,
            "script:language": event.language,
            "script:macro-name": event.macroName,
          })
        : xmlElement("script:event-listener", {
            "script:event-name": event.eventName,
            "script:language": event.language,
            "xlink:type": "simple",
            "xlink:href": event.href,
            "xlink:actuate": event.actuate ?? "onRequest",
          }),
    ),
  );
}

function parseBase(element: Element): OdpSlideNodeBaseOptions {
  const conditional = (
    condition: boolean,
    value: OdpSlideNodeBaseOptions,
  ): OdpSlideNodeBaseOptions => (condition ? value : {});
  return {
    ...(attributeString(element, "xml:id") === undefined
      ? {}
      : { xmlId: attributeString(element, "xml:id") }),
    ...(attributeString(element, "draw:id") === undefined
      ? {}
      : { drawId: attributeString(element, "draw:id") }),
    ...(attributeString(element, "draw:name") === undefined
      ? {}
      : { name: attributeString(element, "draw:name") }),
    ...(attributeNumber(element, "draw:z-index") === undefined
      ? {}
      : { zIndex: attributeNumber(element, "draw:z-index") }),
    ...(attributeString(element, "draw:layer") === undefined
      ? {}
      : { layer: attributeString(element, "draw:layer") }),
    ...conditional(attributeString(element, "draw:style-name") !== undefined, {
      drawStyleName: attributeString(element, "draw:style-name"),
    }),
    ...conditional(attributeString(element, "draw:class-names") !== undefined, {
      drawClassNames: attributeString(element, "draw:class-names"),
    }),
    ...conditional(attributeString(element, "presentation:style-name") !== undefined, {
      presentationStyleName: attributeString(element, "presentation:style-name"),
    }),
    ...conditional(attributeString(element, "presentation:class-names") !== undefined, {
      presentationClassNames: attributeString(element, "presentation:class-names"),
    }),
    ...(attributeString(element, "draw:transform") === undefined &&
    attributeString(element, "dr3d:transform") === undefined
      ? {}
      : {
          transform:
            attributeString(element, "draw:transform") ??
            attributeString(element, "dr3d:transform"),
        }),
    ...(attributeString(element, "draw:caption-id") === undefined
      ? {}
      : { captionId: attributeString(element, "draw:caption-id") }),
    ...(attributeString(element, "draw:text-style-name") === undefined
      ? {}
      : { textStyleName: attributeString(element, "draw:text-style-name") }),
    ...(attributeString(element, "table:end-cell-address") === undefined
      ? {}
      : { tableEndCellAddress: attributeString(element, "table:end-cell-address") }),
    ...(attributeString(element, "table:end-x") === undefined
      ? {}
      : { tableEndX: parseLength(attributeString(element, "table:end-x")!, "/shape") }),
    ...(attributeString(element, "table:end-y") === undefined
      ? {}
      : { tableEndY: parseLength(attributeString(element, "table:end-y")!, "/shape") }),
    ...(attributeString(element, "table:table-background") === undefined
      ? {}
      : { tableBackground: parseBoolean(element, "table:table-background", "/shape") }),
    ...(attributeString(element, "text:anchor-type") === undefined
      ? {}
      : {
          anchorType: parseEnum(
            element,
            "text:anchor-type",
            ["page", "frame", "paragraph", "char", "as-char"],
            "/shape",
          ) as OdpSlideNodeBaseOptions["anchorType"],
        }),
    ...(attributeNumber(element, "text:anchor-page-number") === undefined
      ? {}
      : { anchorPageNumber: attributeNumber(element, "text:anchor-page-number") }),
    ...(attributeString(element, "presentation:class") === undefined
      ? {}
      : { presentationClass: attributeString(element, "presentation:class") }),
    ...(attributeString(element, "presentation:placeholder") === undefined
      ? {}
      : { presentationPlaceholder: parseBoolean(element, "presentation:placeholder", "/shape") }),
    ...(attributeString(element, "presentation:user-transformed") === undefined
      ? {}
      : {
          presentationUserTransformed: parseBoolean(
            element,
            "presentation:user-transformed",
            "/shape",
          ),
        }),
  };
}

function baseAttributes(
  options: OdpSlideNodeBaseOptions,
  transformPrefix: "draw" | "dr3d",
): XmlAttributes {
  return {
    "xml:id": options.xmlId,
    "draw:id": options.drawId,
    "draw:name": options.name,
    "draw:z-index": options.zIndex,
    "draw:layer": options.layer,
    "draw:style-name": options.drawStyleName,
    "draw:class-names": options.drawClassNames,
    "presentation:style-name": options.presentationStyleName,
    "presentation:class-names": options.presentationClassNames,
    [`${transformPrefix}:transform`]: options.transform,
    "draw:caption-id": options.captionId,
    "draw:text-style-name": options.textStyleName,
    "table:end-cell-address": options.tableEndCellAddress,
    "table:end-x": odfLength(options.tableEndX),
    "table:end-y": odfLength(options.tableEndY),
    "table:table-background": options.tableBackground,
    "text:anchor-type": options.anchorType,
    "text:anchor-page-number": options.anchorPageNumber,
    "presentation:class": options.presentationClass,
    "presentation:placeholder": options.presentationPlaceholder,
    "presentation:user-transformed": options.presentationUserTransformed,
  };
}

function shapeAttributeNames(text: boolean): string[] {
  return [
    "draw:caption-id",
    "draw:class-names",
    "draw:id",
    "draw:layer",
    "draw:name",
    "draw:style-name",
    "draw:text-style-name",
    "draw:transform",
    "draw:z-index",
    "draw:page-number",
    "svg:x1",
    "svg:x2",
    "svg:y1",
    "svg:y2",
    "presentation:class",
    "presentation:class-names",
    "presentation:placeholder",
    "presentation:style-name",
    "presentation:user-transformed",
    "svg:height",
    "svg:width",
    "svg:x",
    "svg:y",
    "style:rel-height",
    "style:rel-width",
    "table:end-cell-address",
    "table:end-x",
    "table:end-y",
    "table:table-background",
    "text:anchor-page-number",
    "text:anchor-type",
    "xml:id",
    ...(text ? ["draw:measure-align", "draw:measure-vertical-align"] : []),
  ];
}

function sceneAttributeNames(): string[] {
  return [
    ...shapeAttributeNames(false),
    "dr3d:ambient-color",
    "dr3d:distance",
    "dr3d:focal-length",
    "dr3d:lighting-mode",
    "dr3d:projection",
    "dr3d:shade-mode",
    "dr3d:shadow-slant",
    "dr3d:transform",
    "dr3d:vpn",
    "dr3d:vrp",
    "dr3d:vup",
  ];
}

function parseGeometry(element: Element): Pick<OdpFrameOptions, "x" | "y" | "width" | "height"> {
  const path = `/${element.name}`;
  return {
    ...(attributeString(element, "svg:x") === undefined
      ? {}
      : { x: parseLength(attributeString(element, "svg:x")!, path) }),
    ...(attributeString(element, "svg:y") === undefined
      ? {}
      : { y: parseLength(attributeString(element, "svg:y")!, path) }),
    ...(attributeString(element, "svg:width") === undefined
      ? {}
      : { width: parseLength(attributeString(element, "svg:width")!, path) }),
    ...(attributeString(element, "svg:height") === undefined
      ? {}
      : { height: parseLength(attributeString(element, "svg:height")!, path) }),
  };
}

function parseTitleDescription(element: Element): { title?: string; description?: string } {
  return {
    ...(childNamed(element, "svg:title") === undefined
      ? {}
      : { title: textOf(childNamed(element, "svg:title")) }),
    ...(childNamed(element, "svg:desc") === undefined
      ? {}
      : { description: textOf(childNamed(element, "svg:desc")) }),
  };
}

function parseGluePoints(element: Element): { gluePoints?: OdpGluePointOptions[] } {
  const points = childrenNamed(element, "draw:glue-point");
  if (points.length === 0) return {};
  return {
    gluePoints: points.map((point) => {
      const path = `/${element.name}/draw:glue-point`;
      assertKnown(point, GLUE_POINT_ATTRIBUTES, path);
      const id = attributeNumber(point, "draw:id");
      const escapeDirection = attributeString(point, "draw:escape-direction");
      if (id === undefined) throw requiredAttributeError(path, "draw:id");
      if (!escapeDirection) throw requiredAttributeError(path, "draw:escape-direction");
      return {
        id,
        x: parseDistance(requiredString(point, "svg:x", path), path),
        y: parseDistance(requiredString(point, "svg:y", path), path),
        ...(attributeString(point, "draw:align") === undefined
          ? {}
          : {
              align: parseEnum(
                point,
                "draw:align",
                GLUE_ALIGNMENTS,
                path,
              ) as OdpGluePointOptions["align"],
            }),
        escapeDirection: parseEnum(
          point,
          "draw:escape-direction",
          GLUE_DIRECTIONS,
          path,
        ) as OdpGluePointOptions["escapeDirection"],
      };
    }),
  };
}

function parseLight(element: Element, parentPath: string): OdpLightOptions {
  const path = `${parentPath}/dr3d:light`;
  assertKnown(
    element,
    ["dr3d:diffuse-color", "dr3d:direction", "dr3d:enabled", "dr3d:specular"],
    path,
  );
  const direction = attributeString(element, "dr3d:direction");
  if (!direction) throw requiredAttributeError(path, "dr3d:direction");
  return {
    direction: parseVector(direction, path),
    ...(attributeString(element, "dr3d:diffuse-color") === undefined
      ? {}
      : { diffuseColor: attributeString(element, "dr3d:diffuse-color") }),
    ...(attributeString(element, "dr3d:enabled") === undefined
      ? {}
      : { enabled: parseBoolean(element, "dr3d:enabled", path) }),
    ...(attributeString(element, "dr3d:specular") === undefined
      ? {}
      : { specular: parseBoolean(element, "dr3d:specular", path) }),
  };
}

function parseApplet(element: Element): OdpAppletOptions {
  assertKnown(element, [...APPLET_ATTRIBUTES, ...DATA_ATTRIBUTES], "/draw:frame/draw:applet");
  assertKnownChildren(element, ["draw:param"], "/draw:frame/draw:applet");
  return {
    ...(attributeString(element, "xml:id") === undefined
      ? {}
      : { xmlId: attributeString(element, "xml:id") }),
    ...(attributeString(element, "draw:code") === undefined
      ? {}
      : { code: attributeString(element, "draw:code") }),
    ...(attributeString(element, "draw:object") === undefined
      ? {}
      : { object: attributeString(element, "draw:object") }),
    ...(attributeString(element, "draw:archive") === undefined
      ? {}
      : { archive: attributeString(element, "draw:archive") }),
    ...(attributeString(element, "draw:may-script") === undefined
      ? {}
      : { mayScript: parseBoolean(element, "draw:may-script", "/draw:frame/draw:applet") }),
    ...parseOptionalData(element, "/draw:frame/draw:applet"),
    parameters: childrenNamed(element, "draw:param").map(parseParameter),
  };
}

function parseFloatingFrame(element: Element): OdpFloatingFrameOptions {
  assertKnown(
    element,
    [...FLOATING_FRAME_ATTRIBUTES, ...DATA_ATTRIBUTES],
    "/draw:frame/draw:floating-frame",
  );
  assertKnownChildren(element, [], "/draw:frame/draw:floating-frame");
  return {
    ...(attributeString(element, "xml:id") === undefined
      ? {}
      : { xmlId: attributeString(element, "xml:id") }),
    ...(attributeString(element, "draw:frame-name") === undefined
      ? {}
      : { frameName: attributeString(element, "draw:frame-name") }),
    ...parseRequiredData(element, "/draw:frame/draw:floating-frame"),
  };
}

function parsePlugin(element: Element): OdpPluginOptions {
  assertKnown(element, [...PLUGIN_ATTRIBUTES, ...DATA_ATTRIBUTES], "/draw:frame/draw:plugin");
  assertKnownChildren(element, ["draw:param"], "/draw:frame/draw:plugin");
  return {
    ...(attributeString(element, "xml:id") === undefined
      ? {}
      : { xmlId: attributeString(element, "xml:id") }),
    ...(attributeString(element, "draw:mime-type") === undefined
      ? {}
      : { mimeType: attributeString(element, "draw:mime-type") }),
    ...parseRequiredData(element, "/draw:frame/draw:plugin"),
    parameters: childrenNamed(element, "draw:param").map(parseParameter),
  };
}

function parseTextBox(element: Element): OdpTextBoxOptions {
  const path = "/draw:frame/draw:text-box";
  assertKnown(
    element,
    [
      "draw:auto-grow-height",
      "draw:auto-grow-width",
      "draw:chain-next-name",
      "draw:corner-radius",
      "fo:max-height",
      "fo:max-width",
      "fo:min-height",
      "fo:min-width",
    ],
    path,
  );
  assertKnownChildren(element, ["text:p"], path);
  return {
    ...(attributeString(element, "draw:chain-next-name") === undefined
      ? {}
      : { chainNextName: attributeString(element, "draw:chain-next-name") }),
    ...(attributeString(element, "draw:corner-radius") === undefined
      ? {}
      : { cornerRadius: parseLength(attributeString(element, "draw:corner-radius")!, path) }),
    ...(attributeString(element, "fo:min-height") === undefined
      ? {}
      : { minHeight: parseLength(attributeString(element, "fo:min-height")!, path) }),
    ...(attributeString(element, "fo:max-height") === undefined
      ? {}
      : { maxHeight: parseLength(attributeString(element, "fo:max-height")!, path) }),
    ...(attributeString(element, "fo:min-width") === undefined
      ? {}
      : { minWidth: parseLength(attributeString(element, "fo:min-width")!, path) }),
    ...(attributeString(element, "fo:max-width") === undefined
      ? {}
      : { maxWidth: parseLength(attributeString(element, "fo:max-width")!, path) }),
    ...(attributeString(element, "draw:auto-grow-height") === undefined
      ? {}
      : { autoGrowHeight: parseBoolean(element, "draw:auto-grow-height", path) }),
    ...(attributeString(element, "draw:auto-grow-width") === undefined
      ? {}
      : { autoGrowWidth: parseBoolean(element, "draw:auto-grow-width", path) }),
    text: {
      paragraphs: childrenNamed(element, "text:p").map((paragraph) =>
        parseParagraph(paragraph, new Map()),
      ),
    },
  };
}

function parseTable(element: Element): TableOptions {
  return parseSlideTable(element, new Map(), new Map()).table;
}

function parseParameter(element: Element): OdpParameterOptions {
  assertKnown(element, PARAMETER_ATTRIBUTES, "/draw:param");
  return {
    ...(attributeString(element, "draw:name") === undefined
      ? {}
      : { name: attributeString(element, "draw:name") }),
    ...(attributeString(element, "draw:value") === undefined
      ? {}
      : { value: attributeString(element, "draw:value") }),
  };
}

function parseListeners(element: Element): FormEventListenerOptions[] {
  assertKnownChildren(element, ["script:event-listener"], "/office:event-listeners");
  return (element.elements ?? [])
    .filter((child) => child.type === "element")
    .map((child) => {
      const eventName = attributeString(child, "script:event-name");
      const language = attributeString(child, "script:language");
      if (!eventName || !language) {
        throw new OdpParseError(
          "Event listener requires event name and language",
          "content.xml",
          "/office:event-listeners/script:event-listener",
          "script:event-listener",
          "missing-required-attribute",
        );
      }
      const macroName = attributeString(child, "script:macro-name");
      if (macroName !== undefined) return { target: "macro", eventName, language, macroName };
      const href = attributeString(child, "xlink:href");
      if (!href) {
        throw new OdpParseError(
          "URL event listener requires href",
          "content.xml",
          "/office:event-listeners/script:event-listener",
          "script:event-listener",
          "missing-required-attribute",
        );
      }
      return {
        target: "url",
        eventName,
        language,
        href,
        ...(attributeString(child, "xlink:actuate") === undefined
          ? {}
          : {
              actuate: parseEnum(
                child,
                "xlink:actuate",
                ["onRequest"],
                "/office:event-listeners",
              ) as "onRequest",
            }),
      };
    });
}

function parseOptionalData(element: Element, path: string): { data?: OdpDataReferenceOptions } {
  return attributeString(element, "xlink:href") === undefined
    ? {}
    : { data: parseData(element, path) };
}

function parseRequiredData(element: Element, path: string): { data: OdpDataReferenceOptions } {
  if (attributeString(element, "xlink:href") === undefined)
    throw requiredAttributeError(path, "xlink:href");
  return { data: parseData(element, path) };
}

function parseData(element: Element, path: string): OdpDataReferenceOptions {
  const href = attributeString(element, "xlink:href");
  if (!href) throw requiredAttributeError(path, "xlink:href");
  if (attributeString(element, "xlink:type") !== "simple") {
    throw new OdpParseError(
      "Data reference must be a simple link",
      "content.xml",
      path,
      "xlink:type",
      "invalid-enum",
    );
  }
  const show = attributeString(element, "xlink:show");
  if (show !== undefined && show !== "embed") {
    throw new OdpParseError(
      "Invalid xlink show value",
      "content.xml",
      path,
      "xlink:show",
      "invalid-enum",
    );
  }
  const actuate = attributeString(element, "xlink:actuate");
  if (actuate !== undefined && actuate !== "onLoad") {
    throw new OdpParseError(
      "Invalid xlink actuate value",
      "content.xml",
      path,
      "xlink:actuate",
      "invalid-enum",
    );
  }
  return {
    href,
    ...(attributeString(element, "xlink:show") === undefined ? {} : { show: "embed" as const }),
    ...(attributeString(element, "xlink:actuate") === undefined
      ? {}
      : { actuate: "onLoad" as const }),
  };
}

function parseVector(value: string, path: string): OdpVector3DOptions {
  const parts = value.trim().split(/\s+/u);
  if (parts.length !== 3) {
    throw new OdpParseError(
      "Vector must contain three lengths",
      "content.xml",
      path,
      "vector3D",
      "invalid-vector",
    );
  }
  return {
    x: parseLength(parts[0]!, path),
    y: parseLength(parts[1]!, path),
    z: parseLength(parts[2]!, path),
  };
}

function vectorAttributes(value: OdpVector3DOptions | undefined): string | undefined {
  if (value === undefined) return undefined;
  return [odfLength(value.x), odfLength(value.y), odfLength(value.z)].join(" ");
}

function parseViewBox(element: Element, path: string): [number, number, number, number] {
  const values = (attributeString(element, "svg:viewBox") ?? "").trim().split(/\s+/u).map(Number);
  if (values.length !== 4 || values.some((value) => !Number.isInteger(value))) {
    throw new OdpParseError(
      "Invalid 3D shape viewBox",
      "content.xml",
      path,
      "svg:viewBox",
      "invalid-view-box",
    );
  }
  return [values[0]!, values[1]!, values[2]!, values[3]!];
}

function parseLength(value: string, path: string): Length {
  const emu = lengthToEmu(value);
  if (emu === undefined) {
    throw new OdpParseError(
      `Invalid ODF length: ${value}`,
      "content.xml",
      path,
      value,
      "invalid-length",
    );
  }
  return emu;
}

function odfLength(value: Length | undefined): string | undefined {
  return value === undefined ? undefined : emuToLength(value);
}

function parseDistance(value: string, path: string): Distance {
  if (value === "0") return 0;
  return value.endsWith("%") ? (value as Distance) : parseLength(value, path);
}

function odfDistance(value: Distance): string {
  if (typeof value === "number") return emuToLength(value);
  return value.endsWith("%") ? value : emuToLength(odfLengthToEmu(value));
}

function requiredLength(element: Element, name: string, path: string): Length {
  return parseLength(requiredString(element, name, path), path);
}

function requiredString(element: Element, name: string, path: string): string {
  const value = attributeString(element, name);
  if (value === undefined) throw requiredAttributeError(path, name);
  return value;
}

function parseBoolean(element: Element, name: string, path: string): boolean {
  const value = attributeString(element, name);
  if (value === "true") return true;
  if (value === "false") return false;
  throw new OdpParseError(
    `Invalid boolean for ${name}`,
    "content.xml",
    path,
    name,
    "invalid-boolean",
  );
}

function parseEnum(
  element: Element,
  name: string,
  values: readonly string[],
  path: string,
): string {
  const value = attributeString(element, name) ?? "";
  if (!values.includes(value)) {
    throw new OdpParseError(`Invalid ${name}: ${value}`, "content.xml", path, name, "invalid-enum");
  }
  return value;
}

function assertKnown(element: Element, allowed: readonly string[], path: string): void {
  const names = unknownAttributeNames(element, allowed);
  if (names.length > 0) {
    throw new OdpParseError(
      `Unknown attributes: ${names.join(", ")}`,
      "content.xml",
      path,
      element.name ?? "",
      "unknown-attribute",
    );
  }
}

function assertKnownChildren(element: Element, allowed: readonly string[], path: string): void {
  const child = unknownChildElement(element, allowed);
  if (child) {
    throw new OdpParseError(
      `Unknown child element: ${child.name}`,
      "content.xml",
      `${path}/${child.name}`,
      child.name ?? "",
      "unknown-child-element",
    );
  }
}

function requiredAttributeError(path: string, name: string): OdpParseError {
  return new OdpParseError(
    `${name} is required`,
    "content.xml",
    path,
    name,
    "missing-required-attribute",
  );
}

const GLUE_ALIGNMENTS = [
  "top-left",
  "top",
  "top-right",
  "left",
  "center",
  "right",
  "bottom-left",
  "bottom-right",
] as const;
const GLUE_DIRECTIONS = ["auto", "left", "right", "top", "bottom"] as const;
