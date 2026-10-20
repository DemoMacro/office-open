/**
 * Bridge functions for serializing legacy SlideChild objects to XML.
 *
 * Shared between compiler.ts and descriptor modules (e.g. group.ts)
 * to avoid circular dependencies.
 *
 * @module
 */

import { extUriMatches } from "@office-open/core";
import {
  parseBlackWhiteMode,
  parseNonVisualContentPartProperties,
  parseNonVisualDrawingProperties,
  stringifyNonVisualContentPartProperties,
  stringifyNonVisualDrawingProperties,
  transform2DDesc,
} from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { attr, findChild, findFirst, stringifyElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";

import type { PptxWriteContext } from "../../context";
import { PptxParseError } from "../../pptx-parse-error";
import { chartDesc } from "./chart";
import { groupShapeDesc } from "./group";
import { connectorShapeDesc, lineShapeDesc } from "./line";
import { lockedCanvasDesc } from "./locked-canvas";
import { MEDIA_EXT_URI, audioDesc, videoDesc } from "./media";
import { oleDesc } from "./ole";
import { shapeDesc, pictureDesc } from "./shape";
import { smartArtDesc } from "./smartart";
import { tableDesc } from "./table";

// ── Helpers ──

/**
 * Serialize a SlideChild to XML using descriptors.
 */
export function stringifyChild(child: LegacySlideChild, ctx: PptxWriteContext): string | undefined {
  // JSON object — descriptor dispatch
  if ("shape" in child && child.shape) {
    return shapeDesc.stringify(child.shape, ctx);
  }
  if ("picture" in child && child.picture) {
    return pictureDesc.stringify(child.picture, ctx);
  }
  if ("table" in child && child.table) {
    return tableDesc.stringify(child.table, ctx);
  }
  if ("line" in child && child.line) {
    return lineShapeDesc.stringify(child.line, ctx);
  }
  if ("connector" in child && child.connector) {
    return connectorShapeDesc.stringify(child.connector, ctx);
  }
  if ("group" in child && child.group) {
    return groupShapeDesc.stringify(child.group, ctx);
  }
  if ("chart" in child && child.chart) {
    return chartDesc.stringify(child.chart, ctx);
  }
  if ("smartart" in child && child.smartart) {
    return smartArtDesc.stringify(child.smartart, ctx);
  }
  if ("video" in child && child.video) {
    return videoDesc.stringify(child.video, ctx);
  }
  if ("audio" in child && child.audio) {
    return audioDesc.stringify(child.audio, ctx);
  }
  if ("ole" in child && child.ole) {
    return oleDesc.stringify(child.ole, ctx);
  }
  if ("lockedCanvas" in child && child.lockedCanvas) {
    return lockedCanvasDesc.stringify(child.lockedCanvas, ctx);
  }
  if ("contentPart" in child && child.contentPart) {
    // r:id must name an external relationship the caller owns — without one
    // there is nothing to reference, so the element is skipped
    const part = child.contentPart;
    if (!part.rId) return "";
    const extension =
      part.blackWhiteMode === undefined && !part.nonVisualProperties && !part.transform
        ? ""
        : ' xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';
    const nonVisual = part.nonVisualProperties
      ? `<p14:nvContentPartPr>` +
        stringifyNonVisualDrawingProperties(
          "p14:cNvPr",
          part.nonVisualProperties.drawingProperties.id,
          part.nonVisualProperties.drawingProperties,
          "ContentPart",
        ) +
        stringifyNonVisualContentPartProperties(
          "p14:cNvContentPartPr",
          part.nonVisualProperties.contentPartProperties,
        ) +
        (part.nonVisualProperties.properties?.extensions === undefined
          ? "<p14:nvPr/>"
          : `<p14:nvPr><a:extLst>${part.nonVisualProperties.properties.extensions}</a:extLst></p14:nvPr>`) +
        `</p14:nvContentPartPr>`
      : "";
    const transformXml = part.transform
      ? (transform2DDesc.stringify(part.transform, ctx) ?? "")
          .replace("<a:xfrm", "<p14:xfrm")
          .replace("</a:xfrm>", "</p14:xfrm>")
      : "";
    const blackWhite =
      part.blackWhiteMode === undefined ? "" : ` p14:bwMode="${part.blackWhiteMode}"`;
    return `<p:contentPart${blackWhite}${extension} r:id="${part.rId}">${nonVisual}${transformXml}</p:contentPart>`;
  }
  if ("alternateContent" in child && child.alternateContent) {
    const choices = child.alternateContent.choices.map(
      ({ requires, child: branch }) =>
        `<mc:Choice Requires="${requires}">${stringifyChild(branch, ctx) ?? ""}</mc:Choice>`,
    );
    const fallback =
      child.alternateContent.fallback === undefined
        ? []
        : [
            `<mc:Fallback>${stringifyChild(child.alternateContent.fallback, ctx) ?? ""}</mc:Fallback>`,
          ];
    return `<mc:AlternateContent>${choices.join("")}${fallback.join("")}</mc:AlternateContent>`;
  }
  if ("rawXml" in child && child.rawXml) {
    return child.rawXml;
  }

  return undefined;
}

// ── Parse path ──

/**
 * Parse an XML child element into a SlideChild object using descriptors.
 *
 * Symmetric to {@link stringifyChild} — dispatches based on element tag name
 * to the appropriate descriptor's `parse()` method.
 */
export function parseChild(el: XmlElement, ctx: ReadContext): LegacySlideChild | undefined {
  switch (el.name) {
    case "p:sp": {
      // Check if line shape (prstGeom prst="line")
      const spPr = findChild(el, "p:spPr");
      const prstGeom = spPr ? findChild(spPr, "a:prstGeom") : undefined;
      const prst = prstGeom ? attr(prstGeom, "prst") : undefined;
      if (prst === "line") {
        return { line: lineShapeDesc.parse(el, ctx) } as LegacySlideChild;
      }
      return { shape: shapeDesc.parse(el, ctx) } as LegacySlideChild;
    }
    case "p:pic": {
      const mediaType = detectMediaType(el);
      if (mediaType === "video") {
        return { video: videoDesc.parse(el, ctx) } as LegacySlideChild;
      }
      if (mediaType === "audio") {
        return { audio: audioDesc.parse(el, ctx) } as LegacySlideChild;
      }
      return { picture: pictureDesc.parse(el, ctx) } as LegacySlideChild;
    }
    case "p:graphicFrame":
      return parseGraphicFrameChild(el, ctx);
    case "p:cxnSp":
      return { connector: connectorShapeDesc.parse(el, ctx) } as LegacySlideChild;
    case "p:grpSp":
      return { group: groupShapeDesc.parse(el, ctx) } as LegacySlideChild;
    case "p:contentPart": {
      const rId = attr(el, "r:id");
      if (!rId) return undefined;
      const bwModeKey = Object.keys(el.attributes ?? {}).find(
        (name) => name === "bwMode" || name.endsWith(":bwMode"),
      );
      const nonVisual = findChild(el, "p14:nvContentPartPr");
      const drawingPropertiesElement = nonVisual ? findChild(nonVisual, "p14:cNvPr") : undefined;
      if (nonVisual && !drawingPropertiesElement)
        throw new Error("p14:nvContentPartPr requires p14:cNvPr");
      const contentPartProperties = nonVisual
        ? parseNonVisualContentPartProperties(findChild(nonVisual, "p14:cNvContentPartPr"))
        : undefined;
      const propertiesElement = nonVisual ? findChild(nonVisual, "p14:nvPr") : undefined;
      const transformElement = findChild(el, "p14:xfrm");
      const extensionsElement = propertiesElement
        ? findFirst(propertiesElement, "a:extLst")
        : undefined;
      return {
        contentPart: {
          rId,
          ...(bwModeKey === undefined
            ? {}
            : (() => {
                const blackWhiteMode = parseBlackWhiteMode(el.attributes?.[bwModeKey]);
                if (blackWhiteMode === undefined)
                  throw new Error("p:contentPart has invalid p14:bwMode");
                return { blackWhiteMode };
              })()),
          ...(nonVisual && drawingPropertiesElement
            ? {
                nonVisualProperties: {
                  drawingProperties: {
                    id: attr(drawingPropertiesElement, "id") ?? "",
                    ...parseNonVisualDrawingProperties(drawingPropertiesElement),
                  },
                  ...(contentPartProperties === undefined ? {} : { contentPartProperties }),
                  ...(propertiesElement === undefined
                    ? {}
                    : {
                        properties: extensionsElement
                          ? { extensions: stringifyElement(extensionsElement) }
                          : {},
                      }),
                },
              }
            : {}),
          ...(transformElement ? { transform: transform2DDesc.parse(transformElement, ctx) } : {}),
        },
      } as LegacySlideChild;
    }
    default:
      if (el.name === "mc:AlternateContent") {
        const compatibility = parseCompatibilityChild(el, ctx);
        if (compatibility) return compatibility;
      }
      // Unrecognized element — keep verbatim so round-trip never drops content
      // (vendor extensions, future schema versions).
      return { rawXml: stringifyElement(el) } as LegacySlideChild;
  }
}

/**
 * Select the most broadly representable markup-compatibility branch. Modern
 * branches often contain host-specific controls whose fallback is an ordinary
 * picture; parse candidates first and prefer any branch with a typed model.
 */
function parseCompatibilityChild(el: XmlElement, ctx: ReadContext): LegacySlideChild | undefined {
  const choices = (el.elements ?? [])
    .filter((branch) => branch.name === "mc:Choice")
    .flatMap((branch) => {
      const requires = attr(branch, "Requires");
      if (!requires) return [];
      return (branch.elements ?? [])
        .map((content) => parseChild(content, ctx))
        .filter((child): child is LegacySlideChild => child !== undefined)
        .map((child) => ({ requires, child }));
    });
  const fallbackBranch = findChild(el, "mc:Fallback");
  const fallback = (fallbackBranch?.elements ?? [])
    .map((content) => parseChild(content, ctx))
    .find((child): child is LegacySlideChild => child !== undefined);
  if (!choices.length && fallback === undefined) return undefined;
  return { alternateContent: { choices, ...(fallback ? { fallback } : {}) } };
}

/** Detect whether a p:pic element is a video or audio frame. */
function detectMediaType(el: XmlElement): "video" | "audio" | undefined {
  const nvPicPr = findChild(el, "p:nvPicPr");
  if (!nvPicPr) return undefined;

  const nvPr = findChild(nvPicPr, "p:nvPr");
  if (!nvPr) return undefined;

  // EG_Media elements come first — plain a:videoFile/a:audioFile/etc. without
  // the p14:media extension (pre-2010 files) still identify the frame kind.
  for (const name of ["a:audioCd", "a:wavAudioFile", "a:audioFile"]) {
    if (findChild(nvPr, name)) return "audio";
  }
  for (const name of ["a:videoFile", "a:quickTimeFile"]) {
    if (findChild(nvPr, name)) return "video";
  }

  // Fallback: the p14:media extension. Audio and video frames share the same
  // extension uri, so the frame kind is undecidable here — files without an
  // EG_Media element are near-nonexistent (Office always emits one), and those
  // rare ext-only frames are treated as video.
  const extLst = findChild(nvPr, "p:extLst");
  if (!extLst) return undefined;

  for (const ext of extLst.elements ?? []) {
    if (ext.name !== "p:ext") continue;
    const uri = attr(ext, "uri");
    if (extUriMatches(uri, MEDIA_EXT_URI)) return "video";
  }

  return undefined;
}

/** Dispatch p:graphicFrame to chart/smartart/table descriptor. */
function parseGraphicFrameChild(el: XmlElement, ctx: ReadContext): LegacySlideChild | undefined {
  const graphicData = findFirst(el, "a:graphicData");
  if (!graphicData) return undefined;

  const uri = attr(graphicData, "uri") ?? "";

  if (uri.endsWith("/chart")) {
    const chart = chartDesc.parse(el, ctx);
    if (chart.type === undefined) {
      throw new PptxParseError(
        "Chart graphic frame has no supported chart type",
        "slide",
        "p:graphicFrame/a:graphicData",
        "missing-chart-type",
      );
    }
    return { chart };
  }
  if (uri.includes("/diagram")) {
    return { smartart: smartArtDesc.parse(el, ctx) } as LegacySlideChild;
  }
  if (uri.includes("/ole")) {
    return { ole: oleDesc.parse(el, ctx) } as LegacySlideChild;
  }
  if (uri.endsWith("/lockedCanvas")) {
    return { lockedCanvas: lockedCanvasDesc.parse(el, ctx) } as LegacySlideChild;
  }

  const tbl = findChild(graphicData, "a:tbl");
  if (tbl) {
    return { table: tableDesc.parse(el, ctx) } as LegacySlideChild;
  }

  // Unknown graphicData URI (OLE objects, ink, future content) — keep verbatim.
  return { rawXml: stringifyElement(el) } as LegacySlideChild;
}
