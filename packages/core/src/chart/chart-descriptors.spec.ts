import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../descriptor";
import type { ReadContext, WriteContext } from "../descriptor";
import { chartSpaceDesc } from "./chart-descriptors";
import type { ChartSeriesData, ChartSpaceOptions } from "./types";

function roundTrip(opts: ChartSpaceOptions): ChartSpaceOptions {
  const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parse(chartSpaceDesc, el, {} as ReadContext);
}

describe("chartSpaceDesc", () => {
  it("round-trips column chart type detection", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [{ name: "Revenue", values: [100, 200, 150, 300] }],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("column");
    expect(result.showLegend).not.toBe(false);
  });

  it("round-trips pie chart without legend", () => {
    const opts: ChartSpaceOptions = {
      type: "pie",
      categories: ["A", "B"],
      series: [{ name: "Data", values: [30, 70] }],
      showLegend: false,
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("pie");
    expect(result.showLegend).toBe(false);
  });

  it("round-trips line chart type", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["Jan", "Feb"],
      series: [
        { name: "Series 1", values: [10, 20] },
        { name: "Series 2", values: [30, 40] },
      ],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("line");
  });

  it("round-trips scatter chart type", () => {
    const opts: ChartSpaceOptions = {
      type: "scatter",
      series: [{ name: "Points", values: [1, 2, 3] }],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("scatter");
  });

  it("round-trips bar chart type (distinguished from column via c:barDir)", () => {
    const opts: ChartSpaceOptions = {
      type: "bar",
      style: 2,
      categories: ["X"],
      series: [{ name: "S", values: [1] }],
    };
    const result = roundTrip(opts);
    // bar/column share c:barChart; c:barDir="bar" => horizontal bar
    expect(result.type).toBe("bar");
    expect(result.style).toBe(2);
  });

  it("round-trips column chart type (c:barDir=col)", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["X"],
      series: [{ name: "S", values: [1] }],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("column");
  });

  it("round-trips area chart type", () => {
    const opts: ChartSpaceOptions = {
      type: "area",
      categories: ["A"],
      series: [{ name: "S", values: [1] }],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("area");
  });

  it("round-trips doughnut chart type", () => {
    const opts: ChartSpaceOptions = {
      type: "doughnut",
      categories: ["A", "B"],
      series: [{ name: "S", values: [40, 60] }],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("doughnut");
  });

  it("round-trips radar chart type", () => {
    const opts: ChartSpaceOptions = {
      type: "radar",
      categories: ["A", "B"],
      series: [{ name: "S", values: [10, 20] }],
    };
    const result = roundTrip(opts);
    expect(result.type).toBe("radar");
  });

  it("stringify produces valid XML structure", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      title: "Sales",
      categories: ["Q1"],
      series: [{ name: "R", values: [100] }],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("c:chartSpace");
    expect(xml).toContain("c:chart");
    expect(xml).toContain("c:plotArea");
    expect(xml).toContain("c:barChart");
    expect(xml).toContain("c:ser");
    expect(xml).toContain("Sales");
  });

  // ── Trendlines ──

  it("round-trips series with trendlines", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [
        {
          name: "S1",
          values: [10, 20],
          trendlines: [
            { type: "linear", forward: 2, backward: 1, dispEq: true, dispRSqr: true },
            { type: "poly", order: 3, name: "Poly" },
          ],
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("c:trendline");
    expect(xml).toContain('c:trendlineType val="linear"');
    expect(xml).toContain('c:forward val="2"');
    expect(xml).toContain("<c:dispEq/>");
    expect(xml).toContain("<c:dispRSqr/>");

    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    const trendlines = ser.trendlines ?? [];
    expect(trendlines).toHaveLength(2);
    expect(trendlines[0]?.type).toBe("linear");
    expect(trendlines[0]?.forward).toBe(2);
    expect(trendlines[0]?.dispEq).toBe(true);
    expect(trendlines[1]?.type).toBe("poly");
    expect(trendlines[1]?.order).toBe(3);
    expect(trendlines[1]?.name).toBe("Poly");
  });

  it("round-trips series with exponential trendline period", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A"],
      series: [
        {
          name: "S",
          values: [5],
          trendlines: [{ type: "movingAvg", period: 3 }],
        },
      ],
    };
    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    const trendlines = ser.trendlines ?? [];
    expect(trendlines[0]?.type).toBe("movingAvg");
    expect(trendlines[0]?.period).toBe(3);
  });

  // ── Error bars ──

  it("round-trips series with error bars", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [
        {
          name: "S1",
          values: [10, 20],
          errorBars: {
            direction: "y",
            barType: "both",
            valueType: "fixedVal",
            value: 5,
            noEndCap: false,
          },
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("c:errBars");
    expect(xml).toContain('c:errDir val="y"');
    expect(xml).toContain('c:errValType val="fixedVal"');
    expect(xml).toContain('c:val val="5"');

    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    const errorBars = ser.errorBars;
    expect(errorBars?.direction).toBe("y");
    expect(errorBars?.valueType).toBe("fixedVal");
    expect(errorBars?.value).toBe(5);
    expect(errorBars?.noEndCap).toBe(false);
  });

  it("round-trips error bars with custom plus/minus values", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A"],
      series: [
        {
          name: "S",
          values: [10],
          errorBars: {
            valueType: "cust",
            barType: "both",
            plusValue: 3,
            minusValue: 1.5,
          },
        },
      ],
    };
    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    const errorBars = ser.errorBars;
    expect(errorBars?.plusValue).toBe(3);
    expect(errorBars?.minusValue).toBe(1.5);
  });

  // ── Data labels ──

  it("round-trips series with data labels", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["Q1", "Q2"],
      series: [
        {
          name: "Revenue",
          values: [100, 200],
          dataLabels: {
            position: "outEnd",
            showVal: true,
            showCatName: true,
            showLegendKey: false,
            separator: ", ",
          },
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("c:dLbls");
    expect(xml).toContain('c:dLblPos val="outEnd"');
    expect(xml).toContain("<c:showVal/>");
    expect(xml).toContain('<c:showLegendKey val="0"/>');
    expect(xml).toContain("<c:separator>, </c:separator>");

    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    const dataLabels = ser.dataLabels;
    expect(dataLabels?.position).toBe("outEnd");
    expect(dataLabels?.showVal).toBe(true);
    expect(dataLabels?.showLegendKey).toBe(false);
    expect(dataLabels?.separator).toBe(", ");
  });

  // ── 3D view ──

  it("round-trips view3D settings", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      threeD: true,
      categories: ["A"],
      series: [{ name: "S", values: [1] }],
      view3D: {
        rotX: 30,
        rotY: 20,
        depthPercent: 100,
        rAngAx: true,
        perspective: 30,
      },
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("c:view3D");
    expect(xml).toContain('c:rotX val="30"');
    expect(xml).toContain('c:rotY val="20"');
    expect(xml).toContain("<c:rAngAx/>");

    const result = roundTrip(opts);
    expect(result.view3D).toBeDefined();
    expect(result.view3D?.rotX).toBe(30);
    expect(result.view3D?.rotY).toBe(20);
    expect(result.view3D?.depthPercent).toBe(100);
    expect(result.view3D?.rAngAx).toBe(true);
    expect(result.view3D?.perspective).toBe(30);
  });

  it("emits view3D before plotArea per XSD ordering", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      threeD: true,
      categories: ["A"],
      series: [{ name: "S", values: [1] }],
      view3D: { rotX: 15 },
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toBeDefined();
    const viewIdx = xml!.indexOf("c:view3D");
    const plotIdx = xml!.indexOf("c:plotArea");
    expect(viewIdx).toBeGreaterThan(-1);
    expect(plotIdx).toBeGreaterThan(-1);
    expect(viewIdx).toBeLessThan(plotIdx);
  });

  // ── Phase 1 edge elements: trendlineLbl, per-point dLbl, leaderLines ──

  it("round-trips trendline label number format", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A"],
      series: [
        {
          name: "S",
          values: [1],
          trendlines: [{ type: "linear", label: { numberFormat: "0.00" } }],
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("<c:trendlineLbl>");
    expect(xml).toContain('formatCode="0.00"');

    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    expect(ser.trendlines?.[0]?.label?.numberFormat).toBe("0.00");
  });

  it("round-trips per-point data label overrides (dLbl)", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A", "B", "C"],
      series: [
        {
          name: "S",
          values: [1, 2, 3],
          dataLabels: {
            showVal: true,
            labels: [
              { index: 1, delete: true },
              { index: 2, position: "outEnd", numberFormat: "0.0" },
            ],
          },
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("<c:dLbl>");
    expect(xml).toContain('c:idx val="1"');
    expect(xml).toContain('c:idx val="2"');

    const result = roundTrip(opts);
    const labels = (result.series[0] as ChartSeriesData).dataLabels?.labels ?? [];
    expect(labels).toHaveLength(2);
    expect(labels[0]?.delete).toBe(true);
    expect(labels[1]?.position).toBe("outEnd");
    expect(labels[1]?.numberFormat).toBe("0.0");
  });

  it("round-trips leaderLines element", () => {
    const opts: ChartSpaceOptions = {
      type: "pie",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2], dataLabels: { showPercent: true, leaderLines: true } }],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("<c:leaderLines/>");

    const result = roundTrip(opts);
    expect((result.series[0] as ChartSeriesData).dataLabels?.leaderLines).toBe(true);
  });
});
