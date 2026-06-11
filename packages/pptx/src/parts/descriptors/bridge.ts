/**
 * Bridge functions for serializing legacy SlideChild objects to XML.
 *
 * Shared between compiler.ts and descriptor modules (e.g. group.ts)
 * to avoid circular dependencies.
 *
 * @module
 */

import type { ReadContext } from "@office-open/core/descriptor";
import { attr, findChild, findDeep } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";

import type { PptxWriteContext } from "../../context";
import { chartDesc } from "./chart";
import type { ChartDescriptorOptions } from "./chart";
import { groupShapeDesc } from "./group";
import type { GroupShapeDescriptorOptions } from "./group";
import { connectorShapeDesc } from "./line";
import type { ConnectorShapeDescriptorOptions } from "./line";
import { lineShapeDesc } from "./line";
import type { LineShapeDescriptorOptions } from "./line";
import { lockedCanvasDesc } from "./locked-canvas";
import type { LockedCanvasDescriptorOptions } from "./locked-canvas";
import { audioDesc, videoDesc } from "./media";
import type { AudioDescriptorOptions, VideoDescriptorOptions } from "./media";
import { oleDesc } from "./ole";
import type { OleDescriptorOptions } from "./ole";
import { shapeDesc, pictureDesc } from "./shape";
import type { ShapeDescriptorOptions, PictureDescriptorOptions } from "./shape";
import { smartArtDesc } from "./smartart";
import type { SmartArtDescriptorOptions } from "./smartart";
import { tableDesc } from "./table";
import type { TableDescriptorOptions } from "./table";

// ── Helpers ──

export function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

/**
 * Serialize a SlideChild to XML using descriptors.
 */
export function stringifyChild(child: LegacySlideChild, ctx: PptxWriteContext): string | undefined {
  // JSON object — descriptor dispatch
  if ("shape" in child && child.shape) {
    return shapeDesc.stringify(child.shape as ShapeDescriptorOptions, ctx);
  }
  if ("picture" in child && child.picture) {
    return pictureDesc.stringify(child.picture as PictureDescriptorOptions, ctx);
  }
  if ("table" in child && child.table) {
    return tableDesc.stringify(child.table as TableDescriptorOptions, ctx);
  }
  if ("line" in child && child.line) {
    return lineShapeDesc.stringify(child.line as LineShapeDescriptorOptions, ctx);
  }
  if ("connector" in child && child.connector) {
    return connectorShapeDesc.stringify(child.connector as ConnectorShapeDescriptorOptions, ctx);
  }
  if ("group" in child && child.group) {
    return groupShapeDesc.stringify(child.group as GroupShapeDescriptorOptions, ctx);
  }
  if ("chart" in child && child.chart) {
    return chartDesc.stringify(child.chart as ChartDescriptorOptions, ctx);
  }
  if ("smartart" in child && child.smartart) {
    return smartArtDesc.stringify(child.smartart as SmartArtDescriptorOptions, ctx);
  }
  if ("video" in child && child.video) {
    return videoDesc.stringify(child.video as VideoDescriptorOptions, ctx);
  }
  if ("audio" in child && child.audio) {
    return audioDesc.stringify(child.audio as AudioDescriptorOptions, ctx);
  }
  if ("ole" in child && child.ole) {
    return oleDesc.stringify(child.ole as OleDescriptorOptions, ctx);
  }
  if ("lockedCanvas" in child && child.lockedCanvas) {
    return lockedCanvasDesc.stringify(child.lockedCanvas as LockedCanvasDescriptorOptions, ctx);
  }

  return undefined;
}

// ── Parse path ──

/** Media extension URIs used to detect video/audio in p:pic elements. */
const VIDEO_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
const AUDIO_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";

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
    default:
      return undefined;
  }
}

/** Detect whether a p:pic element is a video or audio frame. */
function detectMediaType(el: XmlElement): "video" | "audio" | undefined {
  const nvPicPr = findChild(el, "p:nvPicPr");
  if (!nvPicPr) return undefined;

  const nvPr = findChild(nvPicPr, "p:nvPr");
  if (!nvPr) return undefined;

  const extLst = findChild(nvPr, "p:extLst");
  if (!extLst) return undefined;

  for (const ext of extLst.elements ?? []) {
    if (ext.name !== "p:ext") continue;
    const uri = attr(ext, "uri");
    if (uri === VIDEO_EXT_URI) return "video";
    if (uri === AUDIO_EXT_URI) return "audio";
  }

  return undefined;
}

/** Dispatch p:graphicFrame to chart/smartart/table descriptor. */
function parseGraphicFrameChild(el: XmlElement, ctx: ReadContext): LegacySlideChild | undefined {
  const graphicData = findDeep(el, "a:graphicData")[0];
  if (!graphicData) return undefined;

  const uri = attr(graphicData, "uri") ?? "";

  if (uri.includes("/chart")) {
    return { chart: chartDesc.parse(el, ctx) } as LegacySlideChild;
  }
  if (uri.includes("/diagram")) {
    return { smartart: smartArtDesc.parse(el, ctx) } as LegacySlideChild;
  }

  const tbl = findChild(graphicData, "a:tbl");
  if (tbl) {
    return { table: tableDesc.parse(el, ctx) } as LegacySlideChild;
  }

  return undefined;
}
