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
import { buildUserShapesData, chartSpaceDesc, userShapesDesc } from "@office-open/core/chart";
import type { ChartSpaceOptions } from "@office-open/core/chart";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringify } from "@office-open/core/descriptor";
import { partPathToRelsPath, resolveRelationshipTarget } from "@office-open/core/opc";
import { themeOverrideDesc } from "@office-open/core/theme";
import { attr, findChild, findFirst } from "@office-open/xml";

import { PptxReadContext, type PptxWriteContext } from "../../context";
import type { ChartOptions } from "../chart-frame";
import {
  readGraphicFrameLocking,
  readGraphicFrameHyperlink,
  stringifyGraphicFrameCnvPr,
  readNvPrPlaceholder,
  stringifyCnvGraphicFramePr,
  stringifyNvPr,
} from "./graphic-frame";
import { readCnvPr, readPositionFromXfrm } from "./shape";
import { nextSlideDrawingId } from "./slide-drawing-ids";

// ── ID counter ──

let _nextChartId = 2048;

/**
 * Separates the graphic-frame identity from the chart payload. Fields such as
 * cNvPr `ext` collide with ChartSpaceOptions `ext`, which owns a different
 * `c:extLst`; forwarding the combined object incorrectly moves frame
 * extensions into the chart part.
 */
export function toChartSpaceOptions(options: ChartOptions): ChartSpaceOptions {
  const {
    title,
    name: _name,
    description: _description,
    hidden: _hidden,
    creationId: _creationId,
    creationIdExtension: _creationIdExtension,
    ext: _extension,
    hyperlink: _hyperlink,
    placeholder: _placeholder,
    placeholderIndex: _placeholderIndex,
    placeholderSize: _placeholderSize,
    placeholderOrientation: _placeholderOrientation,
    hasCustomPrompt: _hasCustomPrompt,
    isPhoto: _isPhoto,
    userDrawn: _userDrawn,
    modId: _modId,
    customerData: _customerData,
    id: _id,
    locking: _locking,
    x: _x,
    y: _y,
    width: _width,
    height: _height,
    chartKey: _chartKey,
    sourcePath: _sourcePath,
    sourceRelationships: _sourceRelationships,
    userShapesSourceTarget: _userShapesSourceTarget,
    themeOverride: _themeOverride,
    ...chartSpace
  } = options;
  return { ...chartSpace, title };
}

// ── Chart descriptor ──

export const chartDesc: CustomDescriptor<ChartOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as PptxWriteContext;
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextChartId++;
    const name = opts.name ?? `Chart ${id}`;
    pptxCtx.registerShapeId(name, id);
    const chartKey = opts.chartKey ?? pptxCtx.nextChartKey();

    // Register chart data with context
    const chartXml = stringify(chartSpaceDesc, toChartSpaceOptions(opts), ctx);
    if (chartXml) {
      pptxCtx.addChart(chartKey, {
        key: chartKey,
        chartSpaceXml: chartXml,
        ...(opts.sourcePath ? { sourcePath: opts.sourcePath } : {}),
        ...(opts.sourceRelationships ? { sourceRelationships: opts.sourceRelationships } : {}),
        ...(opts.userShapes
          ? {
              userShapes: {
                ...buildUserShapesData(opts.userShapes),
                ...(opts.userShapes.path ? { path: opts.userShapes.path } : {}),
                ...(opts.userShapesSourceTarget
                  ? { sourceTarget: opts.userShapesSourceTarget }
                  : {}),
              },
            }
          : {}),
        ...(opts.themeOverride ? { themeOverride: opts.themeOverride } : {}),
      });
    }

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr — strip the chart title so it stays a c:title-only
    // field and never leaks into the cNvPr @title attribute.
    const { title: _chartTitle, ...cNvPrProps } = opts;
    parts.push(
      `<p:nvGraphicFramePr>${stringifyGraphicFrameCnvPr(id, cNvPrProps, name, ctx)}` +
        `${stringifyCnvGraphicFramePr(opts.locking)}` +
        `${stringifyNvPr(opts)}</p:nvGraphicFramePr>`,
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

    // id + name from p:nvGraphicFramePr/p:cNvPr — drop the cNvPr @title so the
    // chart title (parsed from the chart part below) stays the single source.
    const { title: _cNvPrTitle, ...cNvPrProps } = readCnvPr(el, "p:nvGraphicFramePr");
    const locking = readGraphicFrameLocking(findChild(el, "p:nvGraphicFramePr"), _ctx);
    const hyperlink = readGraphicFrameHyperlink(findChild(el, "p:nvGraphicFramePr"), _ctx);
    if (locking !== undefined) result.locking = locking;
    readNvPrPlaceholder(findChild(el, "p:nvGraphicFramePr") ?? el, result);
    Object.assign(result, cNvPrProps);
    if (hyperlink) result.hyperlink = hyperlink;

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
            const chartRelsEl = _ctx.getPart(partPathToRelsPath(chartPath));
            const chartRelsMap = new Map(
              (chartRelsEl?.elements ?? [])
                .filter((element) => element.name === "Relationship")
                .flatMap((element) => {
                  const relationshipId = attr(element, "Id");
                  const target = attr(element, "Target");
                  if (!relationshipId || !target) return [];
                  return [[relationshipId, resolveRelationshipTarget(chartPath, target)]];
                }),
            );
            const chartContext =
              _ctx instanceof PptxReadContext ? _ctx.withRelationships(chartRelsMap) : _ctx;
            Object.assign(result, chartSpaceDesc.parse(chartXml, chartContext));
            // c:userShapes body hangs off the chart part's own rels — the
            // core descriptor reads the r:id only, fill the anchors here
            const relationships =
              chartRelsEl?.elements?.filter((element) => element.name === "Relationship") ?? [];
            const us = result.userShapes;
            const userShapesRel = us
              ? relationships.find(
                  (element) =>
                    attr(element, "Id") === (us.relationshipId ?? "") &&
                    (attr(element, "Type") ?? "").endsWith("/chartUserShapes"),
                )
              : undefined;
            const target = userShapesRel ? attr(userShapesRel, "Target") : undefined;
            const userShapesPath = target
              ? resolveRelationshipTarget(chartPath, target)
              : undefined;
            if (us) {
              const bodyEl = userShapesPath ? _ctx.getPart(userShapesPath) : undefined;
              if (bodyEl) {
                const parsed = userShapesDesc.parse(bodyEl, chartContext);
                us.rootElement = parsed.rootElement;
                us.anchors = parsed.anchors;
              } else if (_ctx.getRaw(userShapesPath ?? "")?.length === 0) us.empty = true;
            }
            if (userShapesPath) {
              const updatedUserShapes = result.userShapes;
              if (updatedUserShapes) updatedUserShapes.path = userShapesPath;
            } else {
              delete result.userShapes;
            }
            if (target) result.userShapesSourceTarget = target;
            const themeOverrideRel = relationships?.find((element) =>
              (attr(element, "Type") ?? "").endsWith("/themeOverride"),
            );
            const themeOverrideTarget = themeOverrideRel
              ? attr(themeOverrideRel, "Target")
              : undefined;
            const themeOverridePath = themeOverrideTarget
              ? resolveRelationshipTarget(chartPath, themeOverrideTarget)
              : undefined;
            const themeOverrideEl = themeOverridePath ? _ctx.getPart(themeOverridePath) : undefined;
            if (themeOverrideRel && themeOverridePath && themeOverrideEl) {
              result.themeOverride = {
                options: themeOverrideDesc.parse(themeOverrideEl, chartContext),
                path: themeOverridePath,
              };
            }
            result.sourcePath = chartPath;
            result.sourceRelationships = relationships
              ?.filter((element) => element !== userShapesRel)
              .flatMap((element) => {
                const relationshipType = attr(element, "Type");
                const relationshipTarget = attr(element, "Target");
                const relationshipId = attr(element, "Id");
                const targetMode = attr(element, "TargetMode");
                if (!relationshipType || !relationshipTarget || !relationshipId) return [];
                return [
                  {
                    relationshipType,
                    target: relationshipTarget,
                    rId: relationshipId,
                    ...(targetMode === "External" ? { targetMode: "External" as const } : {}),
                  },
                ];
              });
          }
        }
      }
    }

    return result as ChartOptions;
  },
};
