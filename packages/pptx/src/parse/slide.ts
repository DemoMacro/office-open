/**
 * Slide parser for PPTX documents.
 *
 * Parses p:sld elements into SlideOptions.
 *
 * @module
 */
import { convertEmuToPixels, xsdLineEndSize } from "@office-open/core";
import {
  attr,
  attrBool,
  attrNum,
  children,
  colorAttr,
  findChild,
  findDeep,
  textOf,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { BackgroundOptions } from "../file/background/background";
import type { ChartOptions } from "../file/chart/chart-frame";
import type { SlideOptions } from "../file/file";
import type { AudioFrameOptions } from "../file/media/audio-frame";
import type { VideoFrameOptions } from "../file/media/video-frame";
import type { PictureOptions } from "../file/picture/picture";
import type { GroupShapeOptions } from "../file/shape/group-shape";
import type {
  ArrowheadType,
  ConnectorShapeOptions,
  LineShapeOptions,
} from "../file/shape/line-shape";
import type { ShapeOptions } from "../file/shape/shape";
import type { SlideChild } from "../file/slide/slide-child";
import type { SmartArtOptions } from "../file/smartart/smartart-frame";
import type { TableOptions } from "../file/table/table-frame";
import type {
  TransitionOptions,
  TransitionDirection,
  TransitionType,
} from "../file/transition/transition";
import { parseTiming } from "./animation";
import type { ParseContext } from "./context";

// ── Slide parser ──────────────────────────────────────────────────────────────

/**
 * Parse a p:sld element into SlideOptions.
 */
export function parseSlide(el: Element, ctx: ParseContext): SlideOptions {
  const opts: Record<string, unknown> = {};

  // p:cSld
  const cSld = findChild(el, "p:cSld");
  if (cSld) {
    // Background
    const bg = findChild(cSld, "p:bg");
    if (bg) opts.background = parseBackground(bg);

    // Shape tree → children
    const spTree = findChild(cSld, "p:spTree");
    if (spTree) {
      const slideChildren: SlideChild[] = [];
      for (const child of spTree.elements ?? []) {
        const parsed = parseSlideChild(child, ctx);
        if (parsed !== undefined) slideChildren.push(parsed);
      }
      if (slideChildren.length > 0) opts.children = slideChildren;
    }

    // Header/Footer (p:hf)
    const hf = findChild(cSld, "p:hf");
    if (hf) {
      const hfOpts: Record<string, unknown> = {};
      if (findChild(hf, "p:sldNum")) hfOpts.slideNumber = true;
      if (findChild(hf, "p:dt")) hfOpts.dateTime = true;
      if (findChild(hf, "p:ftr")) hfOpts.footer = true;
      if (findChild(hf, "p:hdr")) hfOpts.header = true;
      if (Object.keys(hfOpts).length > 0) opts.headerFooter = hfOpts;
    }
  }

  // p:transition
  const transition = findChild(el, "p:transition");
  if (transition) opts.transition = parseTransition(transition);

  // p:timing → animation (attach to shapes by ID)
  const timing = findChild(el, "p:timing");
  if (timing) {
    const animMap = parseTiming(timing);
    if (animMap.size > 0 && opts.children) {
      // Build ID → child index map
      const children = opts.children as SlideChild[];
      for (let i = 0; i < children.length; i++) {
        const child = children[i] as Record<string, unknown>;
        const inner = (child[Object.keys(child)[0]] ?? {}) as Record<string, unknown>;
        const shapeId = inner.id as number | undefined;
        if (shapeId !== undefined && animMap.has(shapeId)) {
          inner.animation = animMap.get(shapeId);
        }
      }
    }
  }

  // Notes via slide rels (notesSlide relationship)
  for (const [, path] of ctx.slideRels) {
    if (!path.includes("notesSlides")) continue;
    const notesEl = ctx.pptx.doc.get(path);
    if (!notesEl) continue;
    const notes = parseNotesSlide(notesEl);
    if (notes) opts.notes = notes;
    break;
  }

  return opts as SlideOptions;
}

// ── Slide child dispatch ──────────────────────────────────────────────────────

export function parseSlideChild(el: Element, ctx: ParseContext): SlideChild | undefined {
  switch (el.name) {
    case "p:sp": {
      // Check if it's a line shape (prstGeom with prst="line")
      const spPr = findChild(el, "p:spPr");
      const prstGeom = spPr ? findChild(spPr, "a:prstGeom") : undefined;
      const prst = prstGeom ? attr(prstGeom, "prst") : undefined;
      if (prst === "line") {
        return { line: parseLineShape(el) };
      }
      return { shape: parseShape(el, ctx) };
    }
    case "p:pic": {
      // Check for video/audio extension URI in nvPr
      const mediaType = detectMediaType(el);
      if (mediaType === "video") {
        return {
          video: parseMediaFrame(el, ctx, "video") as unknown as VideoFrameOptions,
        };
      }
      if (mediaType === "audio") {
        return {
          audio: parseMediaFrame(el, ctx, "audio") as unknown as AudioFrameOptions,
        };
      }
      return { picture: parsePicture(el, ctx) };
    }
    case "p:graphicFrame":
      return parseGraphicFrame(el, ctx);
    case "p:cxnSp":
      return { connector: parseConnector(el) };
    case "p:grpSp":
      return { group: parseGroup(el, ctx) };
    default:
      return undefined;
  }
}

// ── Background parser ─────────────────────────────────────────────────────────

export function parseBackground(el: Element): BackgroundOptions {
  const opts: Record<string, unknown> = {};
  const bgPr = findChild(el, "p:bgPr");
  if (bgPr) {
    const shadeToTitle = attrBool(bgPr, "shadeToTitle");
    if (shadeToTitle) opts.shadeToTitle = true;

    // Parse fill
    const fill = parseFillFromElement(bgPr);
    if (fill) opts.fill = fill;
  }
  return opts as BackgroundOptions;
}

// ── Transition parser ─────────────────────────────────────────────────────────

const DIRECTION_REVERSE: Record<string, TransitionDirection> = {
  l: "left",
  u: "up",
  r: "right",
  d: "down",
  lu: "leftUp",
  ru: "rightUp",
  ld: "leftDown",
  rd: "rightDown",
  out: "out",
  in: "in",
};

const SPEED_MAP: Record<string, "slow" | "med" | "fast"> = {
  slow: "slow",
  med: "med",
  fast: "fast",
};

function parseTransition(el: Element): TransitionOptions {
  const opts: Record<string, unknown> = {};

  const spd = attr(el, "spd");
  if (spd && spd in SPEED_MAP) opts.speed = SPEED_MAP[spd];

  const advClick = attrBool(el, "advClick");
  if (advClick !== undefined) opts.advanceOnClick = advClick;

  const advTm = attrNum(el, "advTm");
  if (advTm !== undefined) opts.advanceAfterTime = advTm;

  // Find transition type element
  for (const child of el.elements ?? []) {
    if (!child.name) continue;
    const typeName = child.name.replace("p:", "") as TransitionType;
    const knownTypes = new Set([
      "fade",
      "push",
      "wipe",
      "split",
      "blinds",
      "checker",
      "comb",
      "randomBar",
      "cover",
      "pull",
      "strips",
      "wheel",
      "zoom",
      "circle",
      "dissolve",
      "diamond",
      "newsflash",
      "plus",
      "wedge",
      "random",
      "cut",
    ]);
    if (knownTypes.has(typeName)) {
      opts.type = typeName;
      const dir = attr(child, "dir");
      if (dir && dir in DIRECTION_REVERSE) opts.direction = DIRECTION_REVERSE[dir];
      const orient = attr(child, "orient");
      if (orient) opts.orient = orient as "horz" | "vert";
      const thruBlk = attrBool(child, "thruBlk");
      if (thruBlk !== undefined) opts.thruBlk = thruBlk;
      const spokes = attrNum(child, "spokes");
      if (spokes !== undefined) opts.spokes = spokes;
      break;
    }
  }

  return opts as TransitionOptions;
}

// ── Shape parser ──────────────────────────────────────────────────────────────

function parseShape(el: Element, ctx: ParseContext): ShapeOptions {
  const opts: Record<string, unknown> = {};

  // p:nvSpPr → id, name, placeholder
  const nvSpPr = findChild(el, "p:nvSpPr");
  if (nvSpPr) {
    const cNvPr = findChild(nvSpPr, "p:cNvPr");
    if (cNvPr) {
      const id = attrNum(cNvPr, "id");
      if (id !== undefined) opts.id = id;
      const name = attr(cNvPr, "name");
      if (name) opts.name = name;
    }
    // Placeholder
    const nvPr = findChild(nvSpPr, "p:nvPr");
    if (nvPr) {
      const ph = findChild(nvPr, "p:ph");
      if (ph) {
        const phType = attr(ph, "type");
        if (phType) opts.placeholder = phType;
        const phIdx = attrNum(ph, "idx");
        if (phIdx !== undefined) opts.placeholderIndex = phIdx;
      }
    }
  }

  // p:spPr → position, geometry, fill, outline, effects
  const spPr = findChild(el, "p:spPr");
  if (spPr) {
    parseTransformFromSpPr(spPr, opts);
    parseGeometryFromSpPr(spPr, opts);
    const fill = parseFillFromElement(spPr);
    if (fill) opts.fill = fill;
    const outline = parseOutlineFromElement(spPr);
    if (outline) opts.outline = outline;
    const effects = parseEffectsFromElement(spPr);
    if (effects) opts.effects = effects;
  }

  // p:txBody → text, paragraphs
  const txBody = findChild(el, "p:txBody");
  if (txBody) {
    parseTextBody(txBody, opts, ctx.slideRels);
  }

  return opts as unknown as ShapeOptions;
}

// ── Picture parser ────────────────────────────────────────────────────────────

// ── Line shape parser ─────────────────────────────────────────────────────────

function parseLineShape(el: Element): LineShapeOptions {
  const opts: Record<string, unknown> = {};

  // p:nvSpPr → id, name
  const nvSpPr = findChild(el, "p:nvSpPr");
  if (nvSpPr) {
    const cNvPr = findChild(nvSpPr, "p:cNvPr");
    if (cNvPr) {
      const id = attrNum(cNvPr, "id");
      if (id !== undefined) opts.id = id;
      const name = attr(cNvPr, "name");
      if (name) opts.name = name;
    }
  }

  // p:spPr → endpoints (same logic as connector)
  const spPr = findChild(el, "p:spPr");
  if (spPr) {
    const xfrm = findChild(spPr, "a:xfrm");
    if (xfrm) {
      const off = findChild(xfrm, "a:off");
      const ext = findChild(xfrm, "a:ext");
      const flipH = attrBool(xfrm, "flipH");
      const flipV = attrBool(xfrm, "flipV");

      if (off && ext) {
        const offX = attrNum(off, "x") ?? 0;
        const offY = attrNum(off, "y") ?? 0;
        const cx = attrNum(ext, "cx") ?? 0;
        const cy = attrNum(ext, "cy") ?? 0;

        const x1 = flipH ? offX + cx : offX;
        const y1 = flipV ? offY + cy : offY;
        const x2 = flipH ? offX : offX + cx;
        const y2 = flipV ? offY : offY + cy;

        opts.x1 = convertEmuToPixels(x1);
        opts.y1 = convertEmuToPixels(y1);
        opts.x2 = convertEmuToPixels(x2);
        opts.y2 = convertEmuToPixels(y2);
      }
    }

    const fill = parseFillFromElement(spPr);
    if (fill) opts.fill = fill;
    const outline = parseOutlineFromElement(spPr);
    if (outline) opts.outline = outline;
  }

  return opts as LineShapeOptions;
}

// ── Picture parser (continued) ────────────────────────────────────────────────

function parsePicture(el: Element, ctx: ParseContext): PictureOptions {
  const opts: Record<string, unknown> = {};

  // Position from p:spPr
  const spPr = findChild(el, "p:spPr");
  if (spPr) parseTransformFromSpPr(spPr, opts);

  // Image data from p:blipFill → a:blip → r:embed
  const blip = findDeep(el, "a:blip")[0];
  if (blip) {
    const rEmbed = attr(blip, "r:embed");
    if (rEmbed) {
      const mediaPath =
        ctx.slideRels.get(rEmbed) ?? ctx.pptx.partRefs.media.find((p) => p.includes(rEmbed));
      if (mediaPath) {
        const imageData = ctx.pptx.doc.getRaw(mediaPath);
        if (imageData) {
          opts.data = imageData;
          opts.type = imageTypeFromPath(mediaPath);
        }
      }
    }
  }

  // Name from p:nvPicPr → a:cNvPr or p:cNvPr
  const nvPicPr = findChild(el, "p:nvPicPr");
  if (nvPicPr) {
    const cNvPr = findChild(nvPicPr, "a:cNvPr") ?? findChild(nvPicPr, "p:cNvPr");
    if (cNvPr) {
      const name = attr(cNvPr, "name");
      if (name) opts.name = name;
    }
  }

  return opts as unknown as PictureOptions;
}

// ── Graphic frame parser (chart / table / smartart) ───────────────────────────

function parseGraphicFrame(el: Element, ctx: ParseContext): SlideChild | undefined {
  const graphicData = findDeep(el, "a:graphicData")[0];
  if (!graphicData) return undefined;

  const uri = attr(graphicData, "uri") ?? "";

  if (uri.includes("/chart")) {
    return { chart: parseChartFrame(el, ctx) };
  }
  if (uri.includes("/diagram")) {
    return { smartart: parseSmartArtFrame(el, ctx) };
  }

  // Check for table (a:graphicData → a:tbl)
  const tbl = findChild(graphicData, "a:tbl");
  if (tbl) {
    return { table: parseTableFrame(el, tbl) };
  }

  return undefined;
}

// ── Chart frame parser ────────────────────────────────────────────────────────

function parseChartFrame(el: Element, ctx: ParseContext): ChartOptions {
  const opts: Record<string, unknown> = {};

  // Position from p:xfrm
  parseTransformFromElement(el, opts);

  // Chart data
  const chartRef = findDeep(el, "c:chart")[0];
  if (chartRef) {
    const rId = attr(chartRef, "r:id");
    if (rId) {
      const chartPath = ctx.slideRels.get(rId);
      if (chartPath) {
        const chartXml = ctx.pptx.doc.get(chartPath);
        if (chartXml) {
          const chartOpts = parseChartXml(chartXml);
          if (chartOpts) Object.assign(opts, chartOpts);
        }
      }
    }
  }

  return opts as unknown as ChartOptions;
}

/**
 * Parse c:chartSpace element into chart options.
 */
function parseChartXml(el: Element): Record<string, unknown> | undefined {
  const chart = findChild(el, "c:chart");
  if (!chart) return undefined;

  const opts: Record<string, unknown> = {};

  // Title
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

  // Series
  const series: { name: string; values: number[] }[] = [];
  let categories: string[] | undefined;

  for (const serEl of typeElement.elements ?? []) {
    if (serEl.name !== "c:ser") continue;
    const nameParts = extractStrCache(serEl, "c:tx");
    const cats = extractStrCache(serEl, "c:cat");
    if (cats.length > 0 && !categories) categories = cats;
    const vals = extractNumCache(serEl);
    series.push({ name: nameParts[0] ?? "", values: vals });
  }

  opts.categories = categories ?? [];
  opts.series = series;
  opts.showLegend = findChild(chart, "c:legend") !== undefined;

  const styleEl = findChild(el, "c:style");
  if (styleEl) {
    const val = attrNum(styleEl, "val");
    if (val !== undefined) opts.style = val;
  }

  return opts;
}

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

// ── SmartArt frame parser ─────────────────────────────────────────────────────

function parseSmartArtFrame(el: Element, ctx: ParseContext): SmartArtOptions {
  const opts: Record<string, unknown> = {};

  // Position
  parseTransformFromElement(el, opts);

  // Name from p:nvGraphicFramePr → p:cNvPr
  const nvGfxFramePr = findChild(el, "p:nvGraphicFramePr");
  if (nvGfxFramePr) {
    const cNvPr = findChild(nvGfxFramePr, "p:cNvPr");
    if (cNvPr) {
      const name = attr(cNvPr, "name");
      if (name) opts.name = name;
    }
  }

  // SmartArt data via dgm:relIds → r:dm
  const relIds = findDeep(el, "dgm:relIds")[0];
  if (relIds) {
    const rId = attr(relIds, "r:dm");
    if (rId) {
      const dataPath = ctx.slideRels.get(rId);
      if (dataPath) {
        const dataEl = ctx.pptx.doc.get(dataPath);
        if (dataEl) {
          parseSmartArtDataXml(dataEl, opts);
        }
      }
    }
  }

  return opts as unknown as SmartArtOptions;
}

function parseSmartArtDataXml(el: Element, opts: Record<string, unknown>): void {
  const ptLst = findChild(el, "dgm:ptLst");
  if (!ptLst) {
    opts.nodes = [];
    return;
  }

  const nodeMap = new Map<string, string>();

  for (const pt of ptLst.elements ?? []) {
    if (pt.name !== "dgm:pt") continue;
    const type = attr(pt, "type");
    const modelId = attr(pt, "modelId");

    if (type === "doc") {
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
      const t = findDeep(pt, "a:t")[0];
      nodeMap.set(modelId, t ? (textOf(t) ?? "") : "");
    }
  }

  // Build tree from connections
  const cxnLst = findChild(el, "dgm:cxnLst");
  if (!cxnLst) {
    opts.nodes = [];
    return;
  }

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

  const topIds = childrenMap.get("0") ?? [];
  opts.nodes = topIds.map((id) => buildSmartArtNode(id, nodeMap, childrenMap));
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

// ── Table frame parser ────────────────────────────────────────────────────────

function parseTableFrame(el: Element, tbl: Element): TableOptions {
  const opts: Record<string, unknown> = {};

  // Position
  parseTransformFromElement(el, opts);

  // a:tblPr
  const tblPr = findChild(tbl, "a:tblPr");
  if (tblPr) {
    if (attrBool(tblPr, "firstRow")) opts.firstRow = true;
    if (attrBool(tblPr, "lastRow")) opts.lastRow = true;
    if (attrBool(tblPr, "bandRow")) opts.bandRow = true;
    if (attrBool(tblPr, "firstCol")) opts.firstCol = true;
    if (attrBool(tblPr, "lastCol")) opts.lastCol = true;
    if (attrBool(tblPr, "bandCol")) opts.bandCol = true;

    // Table-level borders (a:lnL/lnR/lnT/lnB)
    const borders: Record<string, unknown> = {};
    for (const [elName, key] of [
      ["a:lnL", "left"],
      ["a:lnR", "right"],
      ["a:lnT", "top"],
      ["a:lnB", "bottom"],
    ] as const) {
      const borderEl = findChild(tblPr, elName);
      if (borderEl) {
        const borderOpts: Record<string, unknown> = {};
        const w = attrNum(borderEl, "w");
        if (w !== undefined) borderOpts.width = w;
        const fill = parseFillFromElement(borderEl);
        if (typeof fill === "string") borderOpts.color = fill;
        if (Object.keys(borderOpts).length > 0) borders[key] = borderOpts;
      }
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  // a:tblGrid → columnWidths
  const tblGrid = findChild(tbl, "a:tblGrid");
  if (tblGrid) {
    const colWidths: number[] = [];
    for (const gridCol of children(tblGrid, "a:gridCol")) {
      const w = attrNum(gridCol, "w");
      colWidths.push(w ?? 0);
    }
    if (colWidths.length > 0) opts.columnWidths = colWidths;
  }

  // a:tr → rows
  const rows: Record<string, unknown>[] = [];
  for (const tr of children(tbl, "a:tr")) {
    const rowOpts: Record<string, unknown> = {};
    const h = attrNum(tr, "h");
    if (h !== undefined) rowOpts.height = h;

    const cells: Record<string, unknown>[] = [];
    for (const tc of children(tr, "a:tc")) {
      cells.push(parseTableCell(tc));
    }
    rowOpts.cells = cells;
    rows.push(rowOpts);
  }
  opts.rows = rows;

  return opts as unknown as TableOptions;
}

function parseTableCell(tc: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  // gridSpan / rowSpan
  const gridSpan = attrNum(tc, "gridSpan");
  if (gridSpan !== undefined && gridSpan > 1) opts.columnSpan = gridSpan;
  const rowSpan = attrNum(tc, "rowSpan");
  if (rowSpan !== undefined && rowSpan > 1) opts.rowSpan = rowSpan;

  // a:txBody → paragraphs + bodyPr margins
  const txBody = findChild(tc, "a:txBody");
  if (txBody) {
    const paragraphs = parseParagraphsFromTxBody(txBody);
    if (paragraphs.length > 0) opts.children = paragraphs;

    // Extract margins from a:bodyPr (if not already from a:tcPr)
    const bodyPr = findChild(txBody, "a:bodyPr");
    if (bodyPr) {
      const margins: Record<string, number> = {};
      for (const [attrName, key] of [
        ["tIns", "top"],
        ["bIns", "bottom"],
        ["lIns", "left"],
        ["rIns", "right"],
      ] as const) {
        const val = attrNum(bodyPr, attrName);
        if (val !== undefined) margins[key] = val;
      }
      if (Object.keys(margins).length > 0) opts.margins = margins;
    }
  }

  // a:tcPr
  const tcPr = findChild(tc, "a:tcPr");
  if (tcPr) {
    const fill = parseFillFromElement(tcPr);
    if (fill) opts.fill = fill;

    const anchor = attr(tcPr, "anchor");
    if (anchor) {
      const anchorMap: Record<string, string> = { t: "top", ctr: "center", b: "bottom" };
      if (anchor in anchorMap) opts.verticalAlign = anchorMap[anchor];
    }

    // Margins
    const margins: Record<string, number> = {};
    const tIns = attrNum(tcPr, "tIns");
    if (tIns !== undefined) margins.top = tIns;
    const bIns = attrNum(tcPr, "bIns");
    if (bIns !== undefined) margins.bottom = bIns;
    const lIns = attrNum(tcPr, "lIns");
    if (lIns !== undefined) margins.left = lIns;
    const rIns = attrNum(tcPr, "rIns");
    if (rIns !== undefined) margins.right = rIns;
    if (Object.keys(margins).length > 0) opts.margins = margins;

    // Borders
    const borders: Record<string, unknown> = {};
    for (const [elName, key] of [
      ["a:lnL", "left"],
      ["a:lnR", "right"],
      ["a:lnT", "top"],
      ["a:lnB", "bottom"],
    ] as const) {
      const borderEl = findChild(tcPr, elName);
      if (borderEl) {
        const w = attrNum(borderEl, "w");
        const borderOpts: Record<string, unknown> = {};
        if (w !== undefined) borderOpts.width = w;
        const fill = parseFillFromElement(borderEl);
        if (typeof fill === "string") borderOpts.color = fill;
        if (Object.keys(borderOpts).length > 0) borders[key] = borderOpts;
      }
    }
    if (Object.keys(borders).length > 0) opts.borders = borders;
  }

  return opts;
}

// ── Connector parser ──────────────────────────────────────────────────────────

function parseConnector(el: Element): ConnectorShapeOptions {
  const opts: Record<string, unknown> = {};

  // p:nvCxnSpPr → id, name
  const nvCxnSpPr = findChild(el, "p:nvCxnSpPr");
  if (nvCxnSpPr) {
    const cNvPr = findChild(nvCxnSpPr, "p:cNvPr");
    if (cNvPr) {
      const id = attrNum(cNvPr, "id");
      if (id !== undefined) opts.id = id;
      const name = attr(cNvPr, "name");
      if (name) opts.name = name;
    }
  }

  // p:spPr → endpoints
  const spPr = findChild(el, "p:spPr");
  if (spPr) {
    const xfrm = findChild(spPr, "a:xfrm");
    if (xfrm) {
      const off = findChild(xfrm, "a:off");
      const ext = findChild(xfrm, "a:ext");
      const flipH = attrBool(xfrm, "flipH");
      const flipV = attrBool(xfrm, "flipV");

      if (off && ext) {
        const offX = attrNum(off, "x") ?? 0;
        const offY = attrNum(off, "y") ?? 0;
        const cx = attrNum(ext, "cx") ?? 0;
        const cy = attrNum(ext, "cy") ?? 0;

        // Reconstruct endpoints
        const x1 = flipH ? offX + cx : offX;
        const y1 = flipV ? offY + cy : offY;
        const x2 = flipH ? offX : offX + cx;
        const y2 = flipV ? offY : offY + cy;

        opts.x1 = convertEmuToPixels(x1);
        opts.y1 = convertEmuToPixels(y1);
        opts.x2 = convertEmuToPixels(x2);
        opts.y2 = convertEmuToPixels(y2);
      }
    }

    const fill = parseFillFromElement(spPr);
    if (fill) opts.fill = fill;
    const outline = parseOutlineFromElement(spPr);
    if (outline) opts.outline = outline;

    // Arrowheads from outline
    // Note: generation code maps endArrowhead → a:headEnd, beginArrowhead → a:tailEnd
    const ln = findChild(spPr, "a:ln");
    if (ln) {
      const headEnd = findChild(ln, "a:headEnd");
      if (headEnd) {
        const type = attr(headEnd, "type");
        if (type) opts.endArrowhead = type as ArrowheadType;
        const len = attr(headEnd, "len");
        if (len) opts.arrowheadLength = xsdLineEndSize.from(len) as "small" | "medium" | "large";
        const w = attr(headEnd, "w");
        if (w) opts.arrowheadWidth = xsdLineEndSize.from(w) as "small" | "medium" | "large";
      }
      const tailEnd = findChild(ln, "a:tailEnd");
      if (tailEnd) {
        const type = attr(tailEnd, "type");
        if (type) opts.beginArrowhead = type as ArrowheadType;
      }
    }
  }

  return opts as unknown as ConnectorShapeOptions;
}

// ── Group parser ──────────────────────────────────────────────────────────────

function parseGroup(el: Element, ctx: ParseContext): GroupShapeOptions {
  const opts: Record<string, unknown> = {};

  // p:grpSpPr → position, rotation, flip
  const grpSpPr = findChild(el, "p:grpSpPr");
  if (grpSpPr) {
    const xfrm = findChild(grpSpPr, "a:xfrm");
    if (xfrm) {
      const off = findChild(xfrm, "a:off");
      if (off) {
        const x = attrNum(off, "x");
        if (x !== undefined) opts.x = convertEmuToPixels(x);
        const y = attrNum(off, "y");
        if (y !== undefined) opts.y = convertEmuToPixels(y);
      }
      const ext = findChild(xfrm, "a:ext");
      if (ext) {
        const cx = attrNum(ext, "cx");
        if (cx !== undefined) opts.width = convertEmuToPixels(cx);
        const cy = attrNum(ext, "cy");
        if (cy !== undefined) opts.height = convertEmuToPixels(cy);
      }
      const rot = attrNum(xfrm, "rot");
      if (rot !== undefined) opts.rotation = rot;
      const flipH = attrBool(xfrm, "flipH");
      if (flipH) opts.flipHorizontal = true;
    }
  }

  // Children
  const groupChildren: SlideChild[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
    const parsed = parseSlideChild(child, ctx);
    if (parsed !== undefined) groupChildren.push(parsed);
  }
  opts.children = groupChildren;

  return opts as unknown as GroupShapeOptions;
}

// ── Text body / paragraph / run parsers ───────────────────────────────────────

function parseTextBody(
  txBody: Element,
  opts: Record<string, unknown>,
  rels?: Map<string, string>,
): void {
  const textBody: Record<string, unknown> = {};

  // a:bodyPr
  const bodyPr = findChild(txBody, "a:bodyPr");
  if (bodyPr) {
    const vert = attr(bodyPr, "vert");
    if (vert) textBody.vertical = vert;
    const anchor = attr(bodyPr, "anchor");
    if (anchor) {
      const anchorMap: Record<string, string> = { t: "top", ctr: "center", b: "bottom" };
      if (anchor in anchorMap) textBody.anchor = anchorMap[anchor];
    }
    const wrap = attr(bodyPr, "wrap");
    if (wrap) textBody.wrap = wrap;
    const tIns = attrNum(bodyPr, "tIns");
    const bIns = attrNum(bodyPr, "bIns");
    const lIns = attrNum(bodyPr, "lIns");
    const rIns = attrNum(bodyPr, "rIns");
    const margins: Record<string, number> = {};
    if (tIns !== undefined) margins.top = tIns;
    if (bIns !== undefined) margins.bottom = bIns;
    if (lIns !== undefined) margins.left = lIns;
    if (rIns !== undefined) margins.right = rIns;
    if (Object.keys(margins).length > 0) textBody.margins = margins;

    const numCol = attrNum(bodyPr, "numCol");
    if (numCol !== undefined) textBody.columns = numCol;
    const spcCol = attrNum(bodyPr, "spcCol");
    if (spcCol !== undefined) textBody.columnSpacing = spcCol / 100;

    // AutoFit
    if (findChild(bodyPr, "a:normAutofit")) textBody.autoFit = "normal";
    else if (findChild(bodyPr, "a:spAutoFit")) textBody.autoFit = "shape";
    else if (findChild(bodyPr, "a:noAutofit")) textBody.autoFit = "none";
  }

  // Paragraphs
  const paragraphs = parseParagraphsFromTxBody(txBody, rels);
  if (paragraphs.length === 1) {
    const p = paragraphs[0] as Record<string, unknown>;
    // parseParagraph already promotes single-run-plain-text to p.text
    if (p.text && !p.children) {
      const props = p.properties as Record<string, unknown> | undefined;
      // Only promote if paragraph properties are trivial (just bulletNone or empty)
      const hasNonTrivialProps =
        props &&
        Object.keys(props).some(
          (k) => k !== "bullet" || (props[k] as Record<string, unknown>)?.type !== "none",
        );
      if (!hasNonTrivialProps) {
        textBody.text = p.text;
        opts.textBody = textBody;
        return;
      }
    }
  }

  if (paragraphs.length > 0) textBody.children = paragraphs;
  if (Object.keys(textBody).length > 0) opts.textBody = textBody;
}

function parseParagraphsFromTxBody(txBody: Element, rels?: Map<string, string>): unknown[] {
  const paragraphs: unknown[] = [];
  for (const pEl of children(txBody, "a:p")) {
    paragraphs.push(parseParagraph(pEl, rels));
  }
  return paragraphs;
}

function parseParagraph(el: Element, rels?: Map<string, string>): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  // a:pPr
  const pPr = findChild(el, "a:pPr");
  if (pPr) {
    const props: Record<string, unknown> = {};

    const algn = attr(pPr, "algn");
    if (algn) {
      const alignMap: Record<string, string> = {
        l: "left",
        ctr: "center",
        r: "right",
        just: "justify",
      };
      if (algn in alignMap) props.alignment = alignMap[algn];
    }

    const lvl = attrNum(pPr, "lvl");
    if (lvl !== undefined) props.indentLevel = lvl;

    const marL = attrNum(pPr, "marL");
    if (marL !== undefined) props.marginIndent = marL;
    const marR = attrNum(pPr, "marR");
    if (marR !== undefined) props.marginRight = marR;

    // Bullet
    if (findChild(pPr, "a:buNone")) {
      props.bullet = { type: "none" };
    } else if (findChild(pPr, "a:buChar")) {
      const buChar = findChild(pPr, "a:buChar")!;
      const bulletOpts: Record<string, unknown> = { type: "char" };
      const ch = attr(buChar, "char");
      if (ch) bulletOpts.char = ch;
      // Color
      const buClr = findChild(pPr, "a:buClr");
      if (buClr) {
        const srgbClr = findChild(buClr, "a:srgbClr");
        if (srgbClr) {
          const val = colorAttr(srgbClr, "val");
          if (val) bulletOpts.color = val;
        }
      }
      // Size
      const buSzPct = findChild(pPr, "a:buSzPct");
      if (buSzPct) {
        const val = attr(buSzPct, "val");
        if (val) bulletOpts.size = parseInt(String(val), 10);
      }
      props.bullet = bulletOpts;
    } else if (findChild(pPr, "a:buAutoNum")) {
      const buAutoNum = findChild(pPr, "a:buAutoNum")!;
      const bulletOpts: Record<string, unknown> = { type: "autoNum" };
      const fmt = attr(buAutoNum, "type");
      if (fmt) bulletOpts.format = fmt;
      const startAt = attrNum(buAutoNum, "startAt");
      if (startAt !== undefined) bulletOpts.startAt = startAt;
      props.bullet = bulletOpts;
    }

    // Spacing
    const lnSpc = findChild(pPr, "a:lnSpc");
    if (lnSpc) {
      const spcPct = findChild(lnSpc, "a:spcPct");
      if (spcPct) {
        const val = attrNum(spcPct, "val");
        if (val !== undefined) props.lineSpacing = val / 1000;
      }
      const spcPts = findChild(lnSpc, "a:spcPts");
      if (spcPts) {
        const val = attrNum(spcPts, "val");
        if (val !== undefined) props.lineSpacingPoints = val / 100;
      }
    }

    const spcAft = findChild(pPr, "a:spcAft");
    if (spcAft) {
      const spcPts = findChild(spcAft, "a:spcPts");
      if (spcPts) {
        const val = attrNum(spcPts, "val");
        if (val !== undefined) props.marginBottom = val;
      }
    }

    const spcBef = findChild(pPr, "a:spcBef");
    if (spcBef) {
      const spcPts = findChild(spcBef, "a:spcPts");
      if (spcPts) {
        const val = attrNum(spcPts, "val");
        if (val !== undefined) props.marginTop = val;
      }
    }

    if (Object.keys(props).length > 0) opts.properties = props;
  }

  // Runs (a:r)
  const runs: Record<string, unknown>[] = [];
  for (const rEl of children(el, "a:r")) {
    runs.push(parseRun(rEl, rels));
  }

  // Simple text extraction: if single run with only text
  if (runs.length === 1 && Object.keys(runs[0]).length === 1 && runs[0].text) {
    opts.text = runs[0].text;
  } else if (runs.length > 0) {
    opts.children = runs;
  }

  return opts;
}

function parseRun(el: Element, rels?: Map<string, string>): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  // a:rPr
  const rPr = findChild(el, "a:rPr");
  if (rPr) {
    const sz = attrNum(rPr, "sz");
    if (sz !== undefined) opts.fontSize = sz / 100;

    const b = attrBool(rPr, "b");
    if (b !== undefined) opts.bold = b;

    const i = attrBool(rPr, "i");
    if (i !== undefined) opts.italic = i;

    const u = attr(rPr, "u");
    if (u) {
      const underlineMap: Record<string, string> = { sng: "single", dbl: "double", none: "none" };
      if (u in underlineMap) opts.underline = underlineMap[u];
    }

    const lang = attr(rPr, "lang");
    if (lang) opts.lang = lang;

    const strike = attr(rPr, "strike");
    if (strike) {
      const strikeMap: Record<string, string> = {
        sngStrike: "sngStrike",
        dblStrike: "dblStrike",
        noStrike: "noStrike",
      };
      if (strike in strikeMap) opts.strike = strikeMap[strike];
    }

    const baseline = attrNum(rPr, "baseline");
    if (baseline !== undefined) opts.baseline = baseline;

    const cap = attr(rPr, "cap");
    if (cap) {
      const capMap: Record<string, string> = { none: "none", all: "all", small: "small" };
      if (cap in capMap) opts.capitalization = capMap[cap];
    }

    const spc = attrNum(rPr, "spc");
    if (spc !== undefined) opts.spacing = spc;

    // Font
    const latin = findChild(rPr, "a:latin");
    if (latin) {
      const typeface = attr(latin, "typeface");
      if (typeface) opts.font = typeface;
    }

    // Fill (text color)
    const fill = parseFillFromElement(rPr);
    if (fill) opts.fill = fill;

    // Hyperlink (a:hlinkClick)
    const hlinkClick = findChild(rPr, "a:hlinkClick");
    if (hlinkClick && rels) {
      const rId = attr(hlinkClick, "r:id");
      if (rId) {
        const url = rels.get(rId);
        if (url) {
          const hyperlink: Record<string, unknown> = { url };
          const tooltip = attr(hlinkClick, "tooltip");
          if (tooltip) hyperlink.tooltip = tooltip;
          opts.hyperlink = hyperlink;
        }
      }
    }
  }

  // a:t → text
  const t = findChild(el, "a:t");
  if (t) opts.text = textOf(t) ?? "";

  return opts;
}

// ── Common helpers ────────────────────────────────────────────────────────────

function parseTransformFromSpPr(spPr: Element, opts: Record<string, unknown>): void {
  const xfrm = findChild(spPr, "a:xfrm");
  if (!xfrm) return;

  const off = findChild(xfrm, "a:off");
  if (off) {
    const x = attrNum(off, "x");
    if (x !== undefined) opts.x = convertEmuToPixels(x);
    const y = attrNum(off, "y");
    if (y !== undefined) opts.y = convertEmuToPixels(y);
  }

  const ext = findChild(xfrm, "a:ext");
  if (ext) {
    const cx = attrNum(ext, "cx");
    if (cx !== undefined) opts.width = convertEmuToPixels(cx);
    const cy = attrNum(ext, "cy");
    if (cy !== undefined) opts.height = convertEmuToPixels(cy);
  }

  const rot = attrNum(xfrm, "rot");
  if (rot !== undefined) opts.rotation = rot;

  const flipH = attrBool(xfrm, "flipH");
  if (flipH) opts.flipHorizontal = true;
}

function parseTransformFromElement(el: Element, opts: Record<string, unknown>): void {
  // For graphic frames, xfrm may be direct child or under p:xfrm
  const xfrm = findChild(el, "p:xfrm") ?? findDeep(el, "a:xfrm")[0] ?? findDeep(el, "p:xfrm")[0];
  if (!xfrm) return;

  const off = findChild(xfrm, "a:off");
  if (off) {
    const x = attrNum(off, "x");
    if (x !== undefined) opts.x = convertEmuToPixels(x);
    const y = attrNum(off, "y");
    if (y !== undefined) opts.y = convertEmuToPixels(y);
  }

  const ext = findChild(xfrm, "a:ext");
  if (ext) {
    const cx = attrNum(ext, "cx");
    if (cx !== undefined) opts.width = convertEmuToPixels(cx);
    const cy = attrNum(ext, "cy");
    if (cy !== undefined) opts.height = convertEmuToPixels(cy);
  }
}

function parseGeometryFromSpPr(spPr: Element, opts: Record<string, unknown>): void {
  // a:prstGeom → preset
  const prstGeom = findChild(spPr, "a:prstGeom");
  if (prstGeom) {
    const preset = attr(prstGeom, "prst");
    if (preset && preset !== "rect") opts.geometry = preset;
  }
}

function parseFillFromElement(parent: Element): unknown {
  // Solid fill
  const solidFill = findChild(parent, "a:solidFill");
  if (solidFill) {
    const srgbClr = findChild(solidFill, "a:srgbClr");
    if (srgbClr) {
      return colorAttr(srgbClr, "val");
    }
    const schemeClr = findChild(solidFill, "a:schemeClr");
    if (schemeClr) {
      return { type: "scheme", color: attr(schemeClr, "val") };
    }
  }

  // No fill
  if (findChild(parent, "a:noFill")) return { type: "none" };

  // Pattern fill
  const pattFill = findChild(parent, "a:pattFill");
  if (pattFill) {
    const prst = attr(pattFill, "prst");
    const result: Record<string, unknown> = { type: "pattern", pattern: prst ?? "" };

    const fgClr = findChild(pattFill, "a:fgClr");
    if (fgClr) {
      const srgbClr = findChild(fgClr, "a:srgbClr");
      if (srgbClr) {
        const val = colorAttr(srgbClr, "val");
        if (val) result.foregroundColor = val;
      }
    }

    const bgClr = findChild(pattFill, "a:bgClr");
    if (bgClr) {
      const srgbClr = findChild(bgClr, "a:srgbClr");
      if (srgbClr) {
        const val = colorAttr(srgbClr, "val");
        if (val) result.backgroundColor = val;
      }
    }

    return result;
  }

  // Blip fill (image fill on shapes)
  const blipFill = findChild(parent, "a:blipFill");
  if (blipFill) {
    const blip = findChild(blipFill, "a:blip");
    if (blip) {
      return { type: "blip" };
    }
  }

  // Gradient fill
  const gradFill = findChild(parent, "a:gradFill");
  if (gradFill) {
    const stops: { position: number; color: string }[] = [];
    const gsLst = findChild(gradFill, "a:gsLst");
    if (gsLst) {
      for (const gs of gsLst.elements ?? []) {
        if (gs.name !== "a:gs") continue;
        const pos = attrNum(gs, "pos");
        const srgbClr = findChild(gs, "a:srgbClr");
        if (pos !== undefined && srgbClr) {
          stops.push({ position: pos / 1000, color: colorAttr(srgbClr, "val") ?? "" });
        }
      }
    }
    if (stops.length > 0) {
      const result: Record<string, unknown> = { type: "gradient", stops };
      // Extract angle from a:lin element (OOXML unit → degrees)
      const lin = findChild(gradFill, "a:lin");
      if (lin) {
        const ang = attrNum(lin, "ang");
        if (ang !== undefined) result.angle = Math.round(ang / 60000);
      }
      // Extract path shade from a:path element
      const pathEl = findChild(gradFill, "a:path");
      if (pathEl) {
        const pathType = attr(pathEl, "path");
        if (pathType) result.path = pathType;
      }
      return result;
    }
  }

  return undefined;
}

function parseOutlineFromElement(parent: Element): unknown {
  const ln = findChild(parent, "a:ln");
  if (!ln) return undefined;

  const opts: Record<string, unknown> = {};
  const w = attrNum(ln, "w");
  if (w !== undefined) opts.width = w;

  // For outline, only extract solid color (OutlineOptions uses `color` string)
  const solidFill = findChild(ln, "a:solidFill");
  if (solidFill) {
    const srgbClr = findChild(solidFill, "a:srgbClr");
    if (srgbClr) {
      const val = colorAttr(srgbClr, "val");
      if (val) opts.color = val;
    }
  }

  // Dash style
  const prstDash = findChild(ln, "a:prstDash");
  if (prstDash) {
    const val = attr(prstDash, "val");
    if (val) opts.dashStyle = val;
  }

  return Object.keys(opts).length > 0 ? opts : undefined;
}

function parseEffectsFromElement(parent: Element): unknown {
  const spPr = findChild(parent, "p:spPr") ?? parent;
  const opts: Record<string, unknown> = {};

  // 2D effects from a:effectLst
  const effectLst = findChild(spPr, "a:effectLst");
  if (effectLst) {
    for (const child of effectLst.elements ?? []) {
      if (!child.name) continue;
      switch (child.name) {
        case "a:outerShdw": {
          const shadow: Record<string, unknown> = {};
          const blurRad = attrNum(child, "blurRad");
          if (blurRad !== undefined) shadow.blur = blurRad;
          const dist = attrNum(child, "dist");
          if (dist !== undefined) shadow.distance = dist;
          const dir = attrNum(child, "dir");
          if (dir !== undefined) shadow.direction = dir;
          const color = extractColorFromElement(child);
          if (color) {
            if (color.color) shadow.color = color.color;
            if (color.alpha !== undefined) shadow.alpha = color.alpha;
          }
          opts.outerShadow = shadow;
          break;
        }
        case "a:innerShdw": {
          const shadow: Record<string, unknown> = {};
          const blurRad = attrNum(child, "blurRad");
          if (blurRad !== undefined) shadow.blur = blurRad;
          const dist = attrNum(child, "dist");
          if (dist !== undefined) shadow.distance = dist;
          const dir = attrNum(child, "dir");
          if (dir !== undefined) shadow.direction = dir;
          const color = extractColorFromElement(child);
          if (color) {
            if (color.color) shadow.color = color.color;
            if (color.alpha !== undefined) shadow.alpha = color.alpha;
          }
          opts.innerShadow = shadow;
          break;
        }
        case "a:glow": {
          const glow: Record<string, unknown> = {};
          const rad = attrNum(child, "rad");
          if (rad !== undefined) glow.radius = rad;
          const color = extractColorFromElement(child);
          if (color) {
            if (color.color) glow.color = color.color;
            if (color.alpha !== undefined) glow.alpha = color.alpha;
          }
          opts.glow = glow;
          break;
        }
        case "a:reflection": {
          const reflection: Record<string, unknown> = {};
          const blurRad = attrNum(child, "blurRad");
          if (blurRad !== undefined) reflection.blurRadius = blurRad;
          const dist = attrNum(child, "dist");
          if (dist !== undefined) reflection.distance = dist;
          const dir = attrNum(child, "dir");
          if (dir !== undefined) reflection.direction = dir;
          const stA = attrNum(child, "stA");
          if (stA !== undefined) reflection.startAlpha = stA / 1000;
          const endA = attrNum(child, "endA");
          if (endA !== undefined) reflection.endAlpha = endA / 1000;
          opts.reflection = reflection;
          break;
        }
        case "a:softEdge": {
          const rad = attrNum(child, "rad");
          if (rad !== undefined) opts.softEdge = { radius: rad };
          break;
        }
      }
    }
  }

  // 3D scene (rotation3D, lighting) from a:scene3d
  const scene3d = findChild(spPr, "a:scene3d");
  if (scene3d) {
    const camera = findChild(scene3d, "a:camera");
    if (camera) {
      const rot = findChild(camera, "a:rot");
      if (rot) {
        const rotation3d: Record<string, number> = {};
        const lat = attrNum(rot, "lat");
        const lon = attrNum(rot, "lon");
        const rev = attrNum(rot, "rev");
        if (lat !== undefined) rotation3d.x = Math.round(lat / 60000);
        if (lon !== undefined) rotation3d.y = Math.round(lon / 60000);
        if (rev !== undefined) rotation3d.z = Math.round(rev / 60000);
        const fov = attrNum(camera, "fov");
        if (fov !== undefined) rotation3d.perspective = fov;
        if (Object.keys(rotation3d).length > 0) opts.rotation3D = rotation3d;
      }
    }
    const lightRig = findChild(scene3d, "a:lightRig");
    if (lightRig) {
      const rig = attr(lightRig, "rig");
      if (rig) opts.lighting = rig;
    }
  }

  // 3D shape properties (bevel, extrusion, material) from a:sp3d
  const sp3d = findChild(spPr, "a:sp3d");
  if (sp3d) {
    const bevelT = findChild(sp3d, "a:bevelT");
    if (bevelT) {
      const bevel: Record<string, unknown> = {};
      const w = attrNum(bevelT, "w");
      const h = attrNum(bevelT, "h");
      if (w !== undefined) bevel.width = Math.round(w / 12700);
      if (h !== undefined) bevel.height = Math.round(h / 12700);
      const prst = attr(bevelT, "prst");
      if (prst) bevel.preset = prst;
      if (Object.keys(bevel).length > 0) opts.bevelTop = bevel;
    }
    const bevelB = findChild(sp3d, "a:bevelB");
    if (bevelB) {
      const bevel: Record<string, unknown> = {};
      const w = attrNum(bevelB, "w");
      const h = attrNum(bevelB, "h");
      if (w !== undefined) bevel.width = Math.round(w / 12700);
      if (h !== undefined) bevel.height = Math.round(h / 12700);
      const prst = attr(bevelB, "prst");
      if (prst) bevel.preset = prst;
      if (Object.keys(bevel).length > 0) opts.bevelBottom = bevel;
    }
    const z = attrNum(sp3d, "z");
    if (z !== undefined) opts.depth = z;
    const contourW = attrNum(sp3d, "contourW");
    if (contourW !== undefined) opts.contourWidth = contourW;
    const extrusionH = attrNum(sp3d, "extrusionH");
    if (extrusionH !== undefined) opts.extrusionH = extrusionH;
    const prstMaterial = attr(sp3d, "prstMaterial");
    if (prstMaterial) {
      opts.material = prstMaterial;
    }
  }

  return Object.keys(opts).length > 0 ? opts : undefined;
}

/**
 * Extract color and alpha from an effect element's child color element.
 * Looks for a:srgbClr with optional a:alpha transform.
 */
function extractColorFromElement(el: Element): { color?: string; alpha?: number } | undefined {
  for (const child of el.elements ?? []) {
    if (child.name === "a:srgbClr") {
      const color = colorAttr(child, "val");
      let alpha: number | undefined;
      for (const transform of child.elements ?? []) {
        if (transform.name === "a:alpha") {
          const val = attrNum(transform, "val");
          if (val !== undefined) alpha = Math.round(val / 1000);
        }
      }
      return { color: color ?? undefined, alpha };
    }
  }
  return undefined;
}

/**
 * Parse a notesSlide XML element and extract the notes text.
 */
function parseNotesSlide(el: Element): string | undefined {
  const cSld = findChild(el, "p:cSld");
  if (!cSld) return undefined;

  const spTree = findChild(cSld, "p:spTree");
  if (!spTree) return undefined;

  // Look for the notes body placeholder (typically type "body" in p:nvSpPr > p:nvPr > p:ph)
  const texts: string[] = [];
  for (const sp of spTree.elements ?? []) {
    if (sp.name !== "p:sp") continue;
    // Check if this is a body placeholder
    const nvSpPr = findChild(sp, "p:nvSpPr");
    const nvPr = nvSpPr ? findChild(nvSpPr, "p:nvPr") : undefined;
    const ph = nvPr ? findChild(nvPr, "p:ph") : undefined;
    const phType = ph ? attr(ph, "type") : undefined;
    // Skip slide image placeholder (type "sldImg")
    if (phType === "sldImg") continue;

    const txBody = findChild(sp, "p:txBody");
    if (txBody) {
      const t = extractPlainText(txBody);
      if (t) texts.push(t);
    }
  }

  return texts.length > 0 ? texts.join("\n") : undefined;
}

/**
 * Extract plain text from a p:txBody element by concatenating all a:t values.
 */
function extractPlainText(el: Element): string {
  const parts: string[] = [];
  for (const p of children(el, "a:p")) {
    for (const r of children(p, "a:r")) {
      for (const t of children(r, "a:t")) {
        const text = textOf(t);
        if (text) parts.push(text);
      }
    }
    parts.push("\n");
  }
  // Remove trailing newline
  if (parts.length > 0 && parts[parts.length - 1] === "\n") parts.pop();
  return parts.join("");
}

function imageTypeFromPath(path: string): "jpg" | "png" | "gif" | "bmp" | "emf" | "wmf" {
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
    case "emf":
      return "emf";
    case "wmf":
      return "wmf";
    default:
      return "png";
  }
}

// ── Video / Audio detection and parsing ───────────────────────────────────────

const VIDEO_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
const AUDIO_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";

/**
 * Check if a p:pic element is actually a video or audio frame
 * by looking for media extension URIs in nvPr > extLst.
 */
function detectMediaType(el: Element): "video" | "audio" | undefined {
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

function parseMediaFrame(
  el: Element,
  ctx: ParseContext,
  mediaKind: "video" | "audio",
): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  // Position from p:spPr
  const spPr = findChild(el, "p:spPr");
  if (spPr) parseTransformFromSpPr(spPr, opts);

  // Name from p:nvPicPr → a:cNvPr or p:cNvPr
  const nvPicPr = findChild(el, "p:nvPicPr");
  if (nvPicPr) {
    const cNvPr = findChild(nvPicPr, "a:cNvPr") ?? findChild(nvPicPr, "p:cNvPr");
    if (cNvPr) {
      const name = attr(cNvPr, "name");
      if (name) opts.name = name;
    }
  }

  // Media data from a:videoFile/a:audioFile (r:link) or p14:media (r:embed)
  const mediaFileEl = findDeep(el, "a:videoFile")[0] ?? findDeep(el, "a:audioFile")[0];
  const rLink = mediaFileEl ? attr(mediaFileEl, "r:link") : undefined;

  // Fallback: p14:media in extLst (used by audio frames)
  const p14media = !mediaFileEl ? findDeep(el, "p14:media")[0] : undefined;
  const rEmbed = p14media ? attr(p14media, "r:embed") : undefined;

  const mediaRef = rLink ?? rEmbed;
  if (mediaRef) {
    const mediaPath = ctx.slideRels.get(mediaRef);
    if (mediaPath) {
      const data = ctx.pptx.doc.getRaw(mediaPath);
      if (data) {
        opts.data = data;
        opts.type = mediaTypeFromPath(mediaPath, mediaKind);
      }
    }
  }

  // Video poster frame from blipFill
  if (mediaKind === "video") {
    const blipFill = findChild(el, "p:blipFill");
    if (blipFill) {
      const posterBlip = findChild(blipFill, "a:blip");
      if (posterBlip) {
        const rEmbed = attr(posterBlip, "r:embed");
        if (rEmbed) {
          const posterPath = ctx.slideRels.get(rEmbed);
          if (posterPath) {
            const posterData = ctx.pptx.doc.getRaw(posterPath);
            if (posterData) {
              opts.poster = posterData;
              opts.posterType = imageTypeFromPath(posterPath);
            }
          }
        }
      }
    }
  }

  return opts;
}

function mediaTypeFromPath(path: string, kind: "video" | "audio"): string {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";

  if (kind === "video") {
    switch (ext) {
      case "mp4":
        return "mp4";
      case "mov":
      case "qt":
        return "mov";
      case "wmv":
        return "wmv";
      case "avi":
        return "avi";
      default:
        return "mp4";
    }
  } else {
    switch (ext) {
      case "mp3":
        return "mp3";
      case "wav":
        return "wav";
      case "wma":
        return "wma";
      case "aac":
        return "aac";
      default:
        return "mp3";
    }
  }
}

// ── Blip fill parsing ─────────────────────────────────────────────────────────

/**
 * Parse a:blipFill from a shape's spPr.
 * Used when a shape has an image fill (not a picture element).
 * Returns { type: "blip", data } if found.
 */
