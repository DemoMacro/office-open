/**
 * Chart (p:graphicFrame) descriptor for PPTX.
 *
 * Produces a graphicFrame with a chart reference placeholder.
 * The actual chart data is registered in PptxWriteContext for
 * separate compilation by the compiler.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import { chartSpaceDesc } from "@office-open/core/chart";
import type { ChartSpaceOptions } from "@office-open/core/chart";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringify } from "@office-open/core/descriptor";
import {
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
} from "@office-open/core/drawingml";
import { attr, attrNum, findChild, findFirst } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import type { ChartOptions } from "../chart-frame";
import { readPositionFromXfrm } from "./shape";

// ── ID counter ──

let _nextChartId = 2048;

// ── Chart descriptor ──

export const chartDesc: CustomDescriptor<ChartOptions> = {
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

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
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
    const result: Partial<ChartOptions> = {};

    // id + name from p:nvGraphicFramePr/p:cNvPr
    const nvGfxFramePr = findChild(el, "p:nvGraphicFramePr");
    if (nvGfxFramePr) {
      const cNvPr = findChild(nvGfxFramePr, "p:cNvPr");
      if (cNvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    // Position from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Chart data via c:chart → r:id → resolve relationship; the c:chartSpace
    // payload itself is parsed by the core chart descriptor (full type/series
    // coverage — the same descriptor that stringifies it).
    const chartRef = findFirst(el, "c:chart");
    if (chartRef) {
      const rId = attr(chartRef, "r:id");
      if (rId) {
        const chartPath = _ctx.resolveRelationship(rId);
        if (chartPath) {
          const chartXml = _ctx.getPart(chartPath);
          if (chartXml) {
            Object.assign(result, chartSpaceDesc.parse(chartXml, _ctx));
          }
        }
      }
    }

    return result as ChartOptions;
  },
};
