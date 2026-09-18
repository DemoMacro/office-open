import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../../context";
import { parseDrawingRun } from "./drawing-parse";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const CHART_SPACE =
  "<c:chartSpace><c:chart><c:plotArea><c:barChart>" +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit>' +
  '<c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
  "</c:barChart></c:plotArea></c:chart></c:chartSpace>";

function chartContext(): DocxReadContext {
  const chartEl = parseXml(`<root ${NS}>${CHART_SPACE}</root>`).elements?.[0]?.elements?.[0];
  return {
    docx: {
      partRefs: { charts: new Map([["rId1", "word/charts/chart1.xml"]]) },
      doc: {
        get: (path: string) => (path === "word/charts/chart1.xml" ? chartEl : undefined),
      },
    },
  } as unknown as DocxReadContext;
}

function drawingXml(docPrAttrs?: string): string {
  const docPr = docPrAttrs === undefined ? "" : `<wp:docPr ${docPrAttrs}/>`;
  return (
    `<w:drawing ${NS}><wp:inline><wp:extent cx="5486400" cy="3200400"/>` +
    docPr +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart r:id="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
  );
}

function parseDrawing(docPrAttrs?: string) {
  const el = parseXml(drawingXml(docPrAttrs)).elements?.[0];
  if (!el) throw new Error("parsed drawing has no root element");
  return parseDrawingRun(el, chartContext());
}

describe("parseChartDrawing alt text", () => {
  it("carries wp:docPr alt text onto the chart options", () => {
    const result = parseDrawing('id="1" name="Chart 1" descr="Sales chart" title="Sales"');
    expect(result).toMatchObject({
      chart: { altText: { name: "Chart 1", description: "Sales chart", title: "Sales" } },
    });
  });

  it("leaves altText undefined when the drawing carries no wp:docPr", () => {
    const result = parseDrawing();
    const chart = (result as { chart?: { altText?: unknown } } | undefined)?.chart;
    expect(chart?.altText).toBeUndefined();
  });
});
