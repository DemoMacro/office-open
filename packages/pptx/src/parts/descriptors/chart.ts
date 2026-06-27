/**
 * Chart (p:graphicFrame) descriptor for PPTX.
 *
 * Produces a graphicFrame with a chart reference placeholder.
 * The actual chart data is registered in PptxWriteContext for
 * separate compilation by the compiler.
 *
 * @module
 */

import { convertPixelsToEmu } from "@office-open/core";
import { chartSpaceDesc } from "@office-open/core/chart";
import type { ChartSpaceOptions, ChartType } from "@office-open/core/chart";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringify } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, findDeep, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import { readPositionFromXfrm } from "./shape";

// ── Types ──

export interface ChartDescriptorOptions extends ChartSpaceOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Pre-generated chart key (e.g. "chart_2048"). If omitted, auto-generated. */
  chartKey?: string;
}

// ── ID counter ──

let _nextChartId = 2048;

// ── Chart descriptor ──

export const chartDesc: CustomDescriptor<ChartDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextChartId++;
    const name = opts.name ?? `Chart ${id}`;
    const chartKey = opts.chartKey ?? pptxCtx.nextChartKey();

    // Register chart data with context
    const chartXml = stringify(chartSpaceDesc, opts as ChartSpaceOptions, ctx);
    if (chartXml) {
      pptxCtx.addChart(chartKey, { key: chartKey, chartSpaceXml: chartXml });
    }

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 100);
    const h = convertPixelsToEmu(opts.height ?? 100);

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > c:chart (placeholder)
    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{chart:${chartKey}}"/>` +
        `</a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, _ctx) {
    const result: Partial<ChartDescriptorOptions> = {};

    // Position from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Chart data via c:chart → r:id → resolve relationship
    const chartRef = findDeep(el, "c:chart")[0];
    if (chartRef) {
      const rId = attr(chartRef, "r:id");
      if (rId) {
        const chartPath = _ctx.resolveRelationship(rId);
        if (chartPath) {
          const chartXml = _ctx.getPart(chartPath);
          if (chartXml) {
            const chartOpts = parseChartXml(chartXml);
            if (chartOpts) Object.assign(result, chartOpts);
          }
        }
      }
    }

    return result as ChartDescriptorOptions;
  },
};

/** Parse chart XML (c:chartSpace) into chart options. */
function parseChartXml(el: Element): Partial<ChartSpaceOptions> | undefined {
  const chart = findChild(el, "c:chart");
  if (!chart) return undefined;

  const opts: Partial<ChartSpaceOptions> = {};

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
  opts.type = chartType as ChartType;

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

function extractStrCache(serEl: Element, tag: string): string[] {
  const txEl = findChild(serEl, tag);
  if (!txEl) return [];
  const strRef = findChild(txEl, "c:strRef");
  if (!strRef) return [];
  const strCache = findChild(strRef, "c:strCache");
  if (!strCache) return [];

  const parts: string[] = [];
  for (const pt of strCache.elements ?? []) {
    if (pt.name !== "c:pt") continue;
    const v = findChild(pt, "c:v");
    if (v) {
      const text = textOf(v);
      if (text) parts.push(text);
    }
  }
  return parts;
}

function extractNumCache(serEl: Element): number[] {
  const valEl = findChild(serEl, "c:val") ?? findChild(serEl, "c:yVal");
  if (!valEl) return [];
  const numRef = findChild(valEl, "c:numRef");
  if (!numRef) return [];
  const numCache = findChild(numRef, "c:numCache");
  if (!numCache) return [];

  const vals: number[] = [];
  for (const pt of numCache.elements ?? []) {
    if (pt.name !== "c:pt") continue;
    const v = findChild(pt, "c:v");
    if (v) {
      const text = textOf(v);
      if (text) vals.push(parseFloat(text));
    }
  }
  return vals;
}
