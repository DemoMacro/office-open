/**
 * Bridge functions for serializing legacy SlideChild objects to XML.
 *
 * Shared between compiler.ts and descriptor modules (e.g. group.ts)
 * to avoid circular dependencies.
 *
 * @module
 */

import type { ReadContext } from "@office-open/core/descriptor";
import { attr, findChild, findFirst, stringifyElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";

import type { PptxWriteContext } from "../../context";
import { chartDesc } from "./chart";
import { groupShapeDesc } from "./group";
import { connectorShapeDesc, lineShapeDesc } from "./line";
import { lockedCanvasDesc } from "./locked-canvas";
import { audioDesc, videoDesc } from "./media";
import { oleDesc } from "./ole";
import { shapeDesc, pictureDesc } from "./shape";
import { smartArtDesc } from "./smartart";
import { tableDesc } from "./table";

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
  if ("rawXml" in child && child.rawXml) {
    return child.rawXml;
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
      // Unrecognized element — keep verbatim so round-trip never drops content
      // (mc:AlternateContent, vendor extensions, future schema versions).
      return { rawXml: stringifyElement(el) } as LegacySlideChild;
  }
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
  const graphicData = findFirst(el, "a:graphicData");
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

  // Unknown graphicData URI (OLE objects, ink, future content) — keep verbatim.
  return { rawXml: stringifyElement(el) } as LegacySlideChild;
}
