/**
 * Drawing parser for DOCX documents.
 *
 * Parses w:drawing elements and extracts image, chart, or SmartArt data.
 *
 * @module
 */
import { convertEmuToPixels } from "@office-open/core";
import { attr, attrBool, attrNum, findChild, findDeep, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ChartOptions } from "@parts/paragraph/run/chart-run";
import type { ImageOptions } from "@parts/paragraph/run/image-run";
import type { SmartArtOptions } from "@parts/paragraph/run/smartart-run";

import type { DocxReadContext } from "../../context";

/** Union type for parsed drawing child wrappers. */
export type DrawingChild =
  | { image: ImageOptions }
  | { chart: ChartOptions }
  | { smartArt: SmartArtOptions };

/**
 * Parse a w:drawing element and dispatch to the correct parser
 * based on the graphicData URI.
 */
export function parseDrawingRun(el: Element, ctx: DocxReadContext): DrawingChild | undefined {
  const graphicData = findDeep(el, "a:graphicData")[0];
  if (!graphicData) return undefined;

  const uri = attr(graphicData, "uri") ?? "";

  if (uri.includes("/chart")) {
    return parseChartDrawing(el, ctx);
  }
  if (uri.includes("/diagram")) {
    return parseSmartArtDrawing(el, ctx);
  }
  return parseImageRun(el, ctx);
}

/**
 * Determine image type from file extension or MIME type.
 */
function imageTypeFromPath(
  path: string,
): "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf" {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";
  switch (ext) {
    case "jpg":
    case "jpeg":
      return "jpg";
    case "png":
      return "png";
    case "gif":
      return "gif";
    case "bmp":
      return "bmp";
    case "tif":
    case "tiff":
      return "tif";
    case "ico":
      return "ico";
    case "emf":
      return "emf";
    case "wmf":
      return "wmf";
    default:
      return "png"; // fallback
  }
}

/**
 * Parse a w:drawing element and return image data wrapped in { image: ... }.
 */
export function parseImageRun(
  el: Element,
  ctx: DocxReadContext,
): { image: ImageOptions } | undefined {
  // Try inline first, then anchor
  const inline = findDeep(el, "wp:inline")[0];
  const anchor = inline ? undefined : findDeep(el, "wp:anchor")[0];
  const parent = inline ?? anchor;
  if (!parent) return undefined;

  // Get extent (dimensions in EMU, convert to pixels for API)
  const extent = findChild(parent, "wp:extent");
  let width: number | undefined;
  let height: number | undefined;
  if (extent) {
    const cxEmu = attrNum(extent, "cx");
    const cyEmu = attrNum(extent, "cy");
    if (cxEmu !== undefined) width = convertEmuToPixels(cxEmu);
    if (cyEmu !== undefined) height = convertEmuToPixels(cyEmu);
  }

  // Get graphic → graphicData → blip
  const blip = findDeep(parent, "a:blip")[0];
  if (!blip) return undefined;

  const rEmbed = attr(blip, "r:embed");
  if (!rEmbed) return undefined;

  // Look up media path from relationships
  const mediaPath = ctx.docx.partRefs.media.get(rEmbed);
  if (!mediaPath) return undefined;

  // Read image data from ZIP
  const imageData = ctx.docx.doc.getRaw(mediaPath);
  if (!imageData) return undefined;

  const type = imageTypeFromPath(mediaPath);

  const imageOpts: Record<string, unknown> = {
    type,
    data: imageData,
    transformation: {
      ...(width !== undefined ? { width } : {}),
      ...(height !== undefined ? { height } : {}),
    },
  };

  // Get alt text from docPr
  const docPr = findChild(parent, "wp:docPr");
  if (docPr) {
    const name = attr(docPr, "name");
    const descr = attr(docPr, "descr");
    const title = attr(docPr, "title");
    if (name || descr || title) {
      imageOpts.altText = {
        ...(name ? { name } : {}),
        ...(descr ? { description: descr } : {}),
        ...(title ? { title } : {}),
      };
    }
  }

  // For anchor, extract floating properties
  if (anchor && !inline) {
    const floating: Record<string, unknown> = {};
    // Position
    const posH = findChild(anchor, "wp:positionH");
    if (posH) {
      const align = findChild(posH, "wp:align");
      const posOffset = findChild(posH, "wp:posOffset");
      if (align) floating.horizontalPosition = { align: textOf(align) };
      else if (posOffset) {
        const val = Number(textOf(posOffset));
        if (!isNaN(val)) floating.horizontalPosition = { offset: val };
      }
    }
    const posV = findChild(anchor, "wp:positionV");
    if (posV) {
      const align = findChild(posV, "wp:align");
      const posOffset = findChild(posV, "wp:posOffset");
      if (align) floating.verticalPosition = { align: textOf(align) };
      else if (posOffset) {
        const val = Number(textOf(posOffset));
        if (!isNaN(val)) floating.verticalPosition = { offset: val };
      }
    }
    // Wrap
    for (const wrapType of ["wrapSquare", "wrapTight", "wrapTopAndBottom", "wrapNone"]) {
      const wrapEl = findChild(anchor, `wp:${wrapType}`);
      if (wrapEl) {
        floating.wrap = wrapType;
        break;
      }
    }
    // Behind text
    const behindDoc = attrBool(anchor, "behindDoc");
    if (behindDoc) floating.behindDocument = true;

    if (Object.keys(floating).length > 0) {
      imageOpts.floating = floating;
    }
  }

  return { image: imageOpts as unknown as ImageOptions };
}

// ── Common helpers ──────────────────────────────────────────────────────────

function getDrawingExtent(el: Element): { width?: number; height?: number } {
  const inline = findDeep(el, "wp:inline")[0];
  const anchor = inline ? undefined : findDeep(el, "wp:anchor")[0];
  const parent = inline ?? anchor;
  if (!parent) return {};

  const extent = findChild(parent, "wp:extent");
  if (!extent) return {};

  const cxEmu = attrNum(extent, "cx");
  const cyEmu = attrNum(extent, "cy");
  return {
    ...(cxEmu !== undefined ? { width: convertEmuToPixels(cxEmu) } : {}),
    ...(cyEmu !== undefined ? { height: convertEmuToPixels(cyEmu) } : {}),
  };
}

// ── Chart parsing ───────────────────────────────────────────────────────────

/**
 * Look up a relationship ID in a map, with fallback for double "rId" prefix
 * that the library's generation code produces (e.g. "rIdrId7" → "rId7").
 */
function lookupRId(map: Map<string, string>, rId: string | undefined): string | undefined {
  if (!rId) return undefined;
  const direct = map.get(rId);
  if (direct) return direct;
  // Fallback: strip one "rId" prefix when the value starts with "rIdrId"
  if (rId.startsWith("rIdrId")) return map.get(rId.slice(3));
  return undefined;
}

function parseChartDrawing(el: Element, ctx: DocxReadContext): { chart: ChartOptions } | undefined {
  const chartRef = findDeep(el, "c:chart")[0];
  if (!chartRef) return undefined;

  const rId = attr(chartRef, "r:id");
  const chartPath = lookupRId(ctx.docx.partRefs.charts, rId);
  if (!chartPath) return undefined;

  const chartXml = ctx.docx.doc.get(chartPath);
  if (!chartXml) return undefined;

  const opts = parseChartXml(chartXml);
  if (!opts) return undefined;

  const ext = getDrawingExtent(el);
  if (ext.width !== undefined || ext.height !== undefined) {
    (opts as Record<string, unknown>).transformation = {
      ...ext,
    };
  }

  return { chart: opts as unknown as ChartOptions };
}

/**
 * Parse c:chartSpace element into ChartOptions.
 */
function parseChartXml(el: Element): Record<string, unknown> | undefined {
  const chart = findChild(el, "c:chart");
  if (!chart) return undefined;

  const opts: Record<string, unknown> = {};

  // Title: c:chart → c:title → c:tx → c:rich → a:p → a:r → a:t
  const titleEl = findChild(chart, "c:title");
  if (titleEl) {
    const rich = findDeep(titleEl, "c:rich")[0];
    if (rich) {
      const t = findDeep(rich, "a:t")[0];
      if (t) {
        const title = textOf(t);
        if (title) opts.title = title;
      }
    }
  }

  // Plot area → chart type
  const plotArea = findChild(chart, "c:plotArea");
  if (!plotArea) return undefined;

  let chartType: string | undefined;
  let typeElement: Element | undefined;

  for (const child of plotArea.elements ?? []) {
    switch (child.name) {
      case "c:barChart": {
        const barDir = findChild(child, "c:barDir");
        chartType = barDir && attr(barDir, "val") === "bar" ? "bar" : "column";
        typeElement = child;
        break;
      }
      case "c:lineChart":
        chartType = "line";
        typeElement = child;
        break;
      case "c:pieChart":
        chartType = "pie";
        typeElement = child;
        break;
      case "c:areaChart":
        chartType = "area";
        typeElement = child;
        break;
      case "c:scatterChart":
        chartType = "scatter";
        typeElement = child;
        break;
    }
    if (chartType) break;
  }

  if (!chartType || !typeElement) return undefined;
  opts.type = chartType;

  // Parse series
  const series: { name: string; values: number[] }[] = [];
  let categories: string[] | undefined;

  for (const serEl of typeElement.elements ?? []) {
    if (serEl.name !== "c:ser") continue;

    // Series name: c:tx → c:strCache → c:pt → c:v
    const nameParts = extractStrCache(serEl, "c:tx");
    // Categories: c:cat → c:strCache → c:pt → c:v
    const cats = extractStrCache(serEl, "c:cat");
    if (cats.length > 0 && !categories) categories = cats;
    // Values: c:val → c:numCache → c:pt → c:v
    const vals = extractNumCache(serEl);

    series.push({ name: nameParts[0] ?? "", values: vals });
  }

  opts.categories = categories ?? [];
  opts.series = series;

  // Legend
  opts.showLegend = findChild(chart, "c:legend") !== undefined;

  // Style
  const styleEl = findChild(el, "c:style");
  if (styleEl) {
    const val = attrNum(styleEl, "val");
    if (val !== undefined) opts.style = val;
  }

  return opts;
}

/**
 * Extract string values from c:strCache within a container element.
 */
function extractStrCache(parent: Element, containerName: string): string[] {
  const container = findChild(parent, containerName);
  if (!container) return [];
  const cache = findDeep(container, "c:strCache")[0];
  if (!cache) return [];

  const values: string[] = [];
  for (const pt of cache.elements ?? []) {
    if (pt.name !== "c:pt") continue;
    const v = findChild(pt, "c:v");
    if (v) values.push(textOf(v) ?? "");
  }
  return values;
}

/**
 * Extract numeric values from c:numCache within a c:val container.
 */
function extractNumCache(parent: Element): number[] {
  const valEl = findChild(parent, "c:val");
  if (!valEl) return [];
  const cache = findDeep(valEl, "c:numCache")[0];
  if (!cache) return [];

  const values: number[] = [];
  for (const pt of cache.elements ?? []) {
    if (pt.name !== "c:pt") continue;
    const v = findChild(pt, "c:v");
    if (v) {
      const num = Number(textOf(v));
      if (!isNaN(num)) values.push(num);
    }
  }
  return values;
}

// ── SmartArt parsing ────────────────────────────────────────────────────────

function parseSmartArtDrawing(
  el: Element,
  ctx: DocxReadContext,
): { smartArt: SmartArtOptions } | undefined {
  const relIds = findDeep(el, "dgm:relIds")[0];
  if (!relIds) return undefined;

  const rId = attr(relIds, "r:dm");
  const dataPath = lookupRId(ctx.docx.partRefs.diagramData, rId);
  if (!dataPath) return undefined;

  const dataEl = ctx.docx.doc.get(dataPath);
  if (!dataEl) return undefined;

  const opts = parseSmartArtDataXml(dataEl);
  if (!opts) return undefined;

  const ext = getDrawingExtent(el);
  if (ext.width !== undefined || ext.height !== undefined) {
    (opts as Record<string, unknown>).transformation = {
      ...ext,
    };
  }

  return { smartArt: opts as unknown as SmartArtOptions };
}

/**
 * Parse dgm:dataModel element into SmartArtOptions.
 */
function parseSmartArtDataXml(el: Element): Record<string, unknown> | undefined {
  const ptLst = findChild(el, "dgm:ptLst");
  if (!ptLst) return undefined;

  const opts: Record<string, unknown> = {};
  const nodeMap = new Map<string, string>(); // modelId → text

  for (const pt of ptLst.elements ?? []) {
    if (pt.name !== "dgm:pt") continue;
    const type = attr(pt, "type");
    const modelId = attr(pt, "modelId");

    if (type === "doc") {
      // Extract layout/style/color from prSet URIs
      const prSet = findChild(pt, "dgm:prSet");
      if (prSet) {
        const loTypeId = attr(prSet, "loTypeId") ?? "";
        const qsTypeId = attr(prSet, "qsTypeId") ?? "";
        const csTypeId = attr(prSet, "csTypeId") ?? "";

        const layout = loTypeId.split("/").pop();
        if (layout) opts.layout = layout;
        const style = qsTypeId.split("/").pop();
        if (style) opts.style = style;
        const color = csTypeId.split("/").pop();
        if (color) opts.color = color;
      }
    } else if (type === "node" && modelId) {
      // Extract text: dgm:t → a:p → a:r → a:t
      const t = findDeep(pt, "a:t")[0];
      nodeMap.set(modelId, t ? (textOf(t) ?? "") : "");
    }
  }

  // Build tree from connections
  const cxnLst = findChild(el, "dgm:cxnLst");
  if (!cxnLst) {
    opts.data = { nodes: [] };
    return opts;
  }

  // Map: parentId → childIds
  const childrenMap = new Map<string, string[]>();
  for (const cxn of cxnLst.elements ?? []) {
    if (cxn.name !== "dgm:cxn") continue;
    const srcId = attr(cxn, "srcId");
    const destId = attr(cxn, "destId");
    if (!srcId || !destId || !nodeMap.has(destId)) continue;

    let arr = childrenMap.get(srcId);
    if (!arr) {
      arr = [];
      childrenMap.set(srcId, arr);
    }
    arr.push(destId);
  }

  // Root children are connected from doc node (modelId="0")
  const topIds = childrenMap.get("0") ?? [];
  opts.data = { nodes: topIds.map((id) => buildSmartArtNode(id, nodeMap, childrenMap)) };

  return opts;
}

function buildSmartArtNode(
  id: string,
  nodeMap: Map<string, string>,
  childrenMap: Map<string, string[]>,
): { text: string; children?: unknown[] } {
  const text = nodeMap.get(id) ?? "";
  const childIds = childrenMap.get(id) ?? [];

  if (childIds.length === 0) return { text };
  return { text, children: childIds.map((cid) => buildSmartArtNode(cid, nodeMap, childrenMap)) };
}
