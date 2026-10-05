import type { ChartSpaceOptions } from "@office-open/core";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  lengthToEmu,
  xmlElement,
} from "@office-open/odf";
import type { CellOptions, WorksheetChartOptions } from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { unknownOdsElement } from "./workbook";
import { columnNumber } from "./worksheet";

/** A worksheet chart frame collected while serializing an ODS body. */
export interface OdsChartFrame {
  name: string;
  worksheet: string;
  chart: OdsParsedChartOptions;
}

export type OdsParsedChartOptions = WorksheetChartOptions;

/** Worksheet-anchored draw:frame chart objects resolved from the pool. */
export function parseWorksheetCharts(
  body: Element | undefined,
  pool: Map<string, ChartSpaceOptions>,
): OdsChartFrame[] {
  const charts: OdsChartFrame[] = [];
  for (const table of childrenNamed(body, "table:table")) {
    const worksheet = attributeString(table, "table:name");
    const frames: Element[] = [];
    for (const child of table.elements ?? []) {
      if (child.name === "draw:frame") frames.push(child);
      if (child.name === "table:shapes") {
        for (const nested of child.elements ?? [])
          if (nested.name === "draw:frame") frames.push(nested);
      }
    }
    for (const frame of frames) {
      const href = attributeString(childNamed(frame, "draw:object"), "xlink:href")
        ?.replace(/^\.\//, "")
        .replace(/\/$/, "");
      const chart = href ? pool.get(href) : undefined;
      if (!href) continue;
      if (!chart) {
        throw unknownOdsElement(
          frame,
          `/office:spreadsheet/table:table[@table:name="${worksheet ?? ""}"]`,
          href,
          "referenced chart subdocument is missing",
        );
      }
      charts.push({
        name: href,
        worksheet: worksheet ?? "",
        chart: {
          ...chart,
          name: href,
          anchorType: "absolute",
          col: 1,
          row: 1,
          ...endAnchor(frame),
          absoluteX: lengthToEmu(attributeString(frame, "svg:x")),
          absoluteY: lengthToEmu(attributeString(frame, "svg:y")),
          extentCx: lengthToEmu(attributeString(frame, "svg:width")) ?? 400000,
          extentCy: lengthToEmu(attributeString(frame, "svg:height")) ?? 300000,
        },
      });
    }
  }
  return charts;
}

export function endAnchor(frame: Element): Pick<WorksheetChartOptions, "toCol" | "toRow"> {
  const address = attributeString(frame, "table:end-cell-address");
  const match = /^([A-Z]+)([0-9]+)$/.exec(address ?? "");
  if (!match) return {};
  return {
    toCol: columnNumber(match[1]!),
    toRow: Number(match[2]),
  };
}

export function cellGraphicXml(graphic: NonNullable<CellOptions["graphics"]>[number]): string {
  const common = {
    "draw:name": graphic.name,
    "svg:x": emuToLength(graphic.x),
    "svg:y": emuToLength(graphic.y),
    "svg:width": emuToLength(graphic.width),
    "svg:height": emuToLength(graphic.height),
    "draw:z-index": graphic.zIndex,
    "table:end-cell-address": graphic.endCellAddress,
    ...(graphic.endX !== undefined ? { "table:end-x": emuToLength(graphic.endX) } : {}),
    ...(graphic.endY !== undefined ? { "table:end-y": emuToLength(graphic.endY) } : {}),
    "draw:id": graphic.id,
    "draw:style-name": graphic.styleName,
    "draw:text-style-name": graphic.textStyleName,
    "style:rel-width": graphic.relativeWidth,
    "style:rel-height": graphic.relativeHeight,
  };
  const link = {
    "xlink:href": graphic.href,
    "xlink:type": "simple",
    "xlink:show": "embed",
    "xlink:actuate": "onLoad",
  };
  return xmlElement("draw:frame", common, [
    xmlElement(graphic.type === "image" ? "draw:image" : "draw:object", link),
    ...(graphic.description
      ? [xmlElement("svg:desc", undefined, [escapeText(graphic.description)])]
      : []),
  ]);
}
