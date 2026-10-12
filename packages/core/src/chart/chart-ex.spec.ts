import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ReadContext, WriteContext } from "../descriptor";
import { parse as parseDescriptor, stringify as stringifyDescriptor } from "../descriptor";
import { chartExSpaceDesc } from "./chart-ex-descriptor";

const READ_CONTEXT = {} as ReadContext;
const WRITE_CONTEXT = {} as WriteContext;
const root = (xml: string) => {
  const element = parse(xml).elements?.[0];
  if (!element) throw new Error("test fixture has no root element");
  return element;
};

const CHART_EX_XML = `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cx:chartData><cx:externalData r:id="rId1" cx:autoUpdate="0"/><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">12</cx:pt><cx:pt idx="1">-4</cx:pt><cx:pt idx="2">30</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"/><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall" uniqueId="urn:coverage:series"><cx:tx><cx:txData><cx:f>Sheet1!$A$1</cx:f><cx:v>Series</cx:v></cx:txData></cx:tx><cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="1"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea><cx:legend pos="t" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>`;

describe("ChartEx descriptor", () => {
  it("projects a corpus-shaped waterfall into canonical options", () => {
    const options = parseDescriptor(chartExSpaceDesc, root(CHART_EX_XML), READ_CONTEXT);

    expect(options.chartData?.externalData).toEqual({
      relationshipId: "rId1",
      autoUpdate: false,
      autoUpdatePrefix: "cx",
    });
    expect(options.chartData?.data[0]).toMatchObject({
      id: 0,
      dimensions: [
        {
          type: "val",
          formula: { text: "Sheet1!$A$2:$A$4" },
          levels: [
            {
              pointCount: 3,
              formatCode: "General",
              points: [
                { index: 0, value: 12 },
                { index: 1, value: -4 },
                { index: 2, value: 30 },
              ],
            },
          ],
        },
      ],
    });
    expect(options.chart).toMatchObject({
      title: { position: "top", alignment: "center", overlay: false },
      plotArea: {
        region: {
          series: [
            {
              layout: "waterfall",
              text: { formula: { text: "Sheet1!$A$1" }, value: "Series" },
              dataLabels: { position: "outsideEnd", visibility: { value: true } },
              dataId: 0,
              layoutProperties: { subtotals: [1] },
            },
          ],
        },
        axes: [
          { id: 0, scaling: { kind: "category", gapWidth: 0.5 } },
          { id: 1, scaling: { kind: "value" } },
        ],
      },
      legend: { position: "top", alignment: "center", overlay: false },
    });
  });

  it("reconstructs the same model from generated XML", () => {
    const source = root(CHART_EX_XML);
    const options = parseDescriptor(chartExSpaceDesc, source, READ_CONTEXT);
    const generated = stringifyDescriptor(chartExSpaceDesc, options, WRITE_CONTEXT)!;
    const reparsed = parseDescriptor(chartExSpaceDesc, root(generated), READ_CONTEXT);

    expect(generated).toContain('<cx:series layoutId="waterfall"');
    expect(generated).toContain('cx:autoUpdate="0"');
    expect(generated).toContain("<cx:plotAreaRegion>");
    expect(generated).toContain('pos="outEnd"');
    expect(reparsed).toEqual(options);
  });

  it("rejects malformed series and roots with structured failures", () => {
    expect(() => parseDescriptor(chartExSpaceDesc, root("<chart/>"), READ_CONTEXT)).toThrow(
      "ChartEx root must be chartSpace",
    );
    expect(() =>
      stringifyDescriptor(chartExSpaceDesc, { chartData: { data: [] } }, WRITE_CONTEXT),
    ).toThrow("requires chart or dashboard");
  });
});
