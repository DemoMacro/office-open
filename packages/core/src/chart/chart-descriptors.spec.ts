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

  // ── Axes (EG_AxShared + CT_CatAx/CT_ValAx/CT_DateAx/CT_SerAx) ──

  it("emits and parses back default axes for a column chart", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A"],
      series: [{ name: "S", values: [1] }],
    };
    const result = roundTrip(opts);
    expect(result.axes).toBeDefined();
    expect(result.axes).toHaveLength(2);
    expect(result.axes?.[0]?.kind).toBe("category");
    expect(result.axes?.[1]?.kind).toBe("value");
  });

  it("round-trips default axes byte-stably (parse re-emits identical XML)", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
    };
    const xml1 = stringify(chartSpaceDesc, opts, {} as WriteContext);
    const result = roundTrip(opts);
    const xml2 = stringify(chartSpaceDesc, result, {} as WriteContext);
    expect(xml2).toBe(xml1);
  });

  it("round-trips custom value axis (gridlines/units/tickMarks/dispUnits/scaling/crossesAt)", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      axes: [
        {
          kind: "category",
          id: 10,
          crossAxisId: 20,
          position: "b",
          crosses: "autoZero",
          majorGridlines: true,
          labelOffset: 100,
        },
        {
          kind: "value",
          id: 20,
          crossAxisId: 10,
          position: "l",
          numberFormat: "0.0",
          majorUnit: 5,
          minorUnit: 1,
          majorTickMark: "out",
          minorTickMark: "in",
          tickLabelPosition: "low",
          crossBetween: "between",
          crossesAt: 0,
          majorGridlines: true,
          minorGridlines: true,
          scaling: { orientation: "minMax", min: 0, max: 10, logBase: 10 },
          displayUnits: { builtInUnit: "thousands", label: true },
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("<c:majorGridlines/>");
    expect(xml).toContain("<c:minorGridlines/>");
    expect(xml).toContain('c:majorUnit val="5"');
    expect(xml).toContain('c:minorUnit val="1"');
    expect(xml).toContain('c:majorTickMark val="out"');
    expect(xml).toContain('c:tickLblPos val="low"');
    expect(xml).toContain('c:crossBetween val="between"');
    expect(xml).toContain('c:crossesAt val="0"');
    expect(xml).toContain('c:logBase val="10"');
    expect(xml).toContain("<c:dispUnits>");
    expect(xml).toContain('c:builtInUnit val="thousands"');
    expect(xml).toContain("<c:dispUnitsLbl/>");

    const result = roundTrip(opts);
    const value = result.axes?.find((a) => a.kind === "value");
    expect(value?.majorUnit).toBe(5);
    expect(value?.minorUnit).toBe(1);
    expect(value?.majorTickMark).toBe("out");
    expect(value?.minorTickMark).toBe("in");
    expect(value?.tickLabelPosition).toBe("low");
    expect(value?.crossBetween).toBe("between");
    expect(value?.crossesAt).toBe(0);
    expect(value?.scaling?.logBase).toBe(10);
    expect(value?.scaling?.min).toBe(0);
    expect(value?.scaling?.max).toBe(10);
    expect(value?.displayUnits?.builtInUnit).toBe("thousands");
    expect(value?.displayUnits?.label).toBe(true);
    const category = result.axes?.find((a) => a.kind === "category");
    expect(category?.majorGridlines).toBe(true);
  });

  it("round-trips a date axis (baseTimeUnit/units/timeUnits)", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A"],
      series: [{ name: "S", values: [1] }],
      axes: [
        {
          kind: "date",
          id: 10,
          crossAxisId: 20,
          position: "b",
          crosses: "autoZero",
          baseTimeUnit: "days",
          majorUnit: 30,
          majorTimeUnit: "months",
          minorUnit: 7,
          minorTimeUnit: "days",
        },
        { kind: "value", id: 20, crossAxisId: 10, position: "l", crosses: "autoZero" },
      ],
    };
    const result = roundTrip(opts);
    const date = result.axes?.find((a) => a.kind === "date");
    expect(date?.baseTimeUnit).toBe("days");
    expect(date?.majorUnit).toBe(30);
    expect(date?.majorTimeUnit).toBe("months");
    expect(date?.minorUnit).toBe(7);
    expect(date?.minorTimeUnit).toBe("days");
  });

  it("emits custom display unit via custUnit", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A"],
      series: [{ name: "S", values: [1] }],
      axes: [
        { kind: "category", id: 10, crossAxisId: 20, position: "b", crosses: "autoZero" },
        {
          kind: "value",
          id: 20,
          crossAxisId: 10,
          position: "l",
          crosses: "autoZero",
          displayUnits: { customUnit: 50 },
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('c:custUnit val="50"');
  });

  // ── Series enhancements (marker/dPt/invertIfNegative/smooth/explosion/pictureOptions/shape) ──

  it("round-trips line series marker and smooth", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2], marker: { symbol: "circle", size: 7 }, smooth: true }],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('c:symbol val="circle"');
    expect(xml).toContain('c:size val="7"');
    expect(xml).toContain("<c:smooth/>");

    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    expect(ser.marker?.symbol).toBe("circle");
    expect(ser.marker?.size).toBe(7);
    expect(ser.smooth).toBe(true);
  });

  it("round-trips 3D bar series invertIfNegative, pictureOptions, and shape", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      threeD: true,
      categories: ["A", "B"],
      series: [
        {
          name: "S",
          values: [1, -2],
          invertIfNegative: true,
          pictureOptions: {
            applyToFront: true,
            pictureFormat: "stretch",
            pictureStackUnit: 5,
          },
          shape: "cylinder",
        },
      ],
    };
    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    expect(ser.invertIfNegative).toBe(true);
    expect(ser.pictureOptions?.applyToFront).toBe(true);
    expect(ser.pictureOptions?.pictureFormat).toBe("stretch");
    expect(ser.pictureOptions?.pictureStackUnit).toBe(5);
    expect(ser.shape).toBe("cylinder");
  });

  it("round-trips pie series explosion and per-point data overrides", () => {
    const opts: ChartSpaceOptions = {
      type: "pie",
      categories: ["A", "B", "C"],
      series: [
        {
          name: "S",
          values: [1, 2, 3],
          explosion: 10,
          dataPoints: [
            { index: 0, explosion: 30 },
            { index: 1, invertIfNegative: false },
          ],
        },
      ],
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('c:explosion val="10"');
    expect(xml).toContain("<c:dPt>");
    expect(xml).toContain('c:idx val="0"');

    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    expect(ser.explosion).toBe(10);
    expect(ser.dataPoints).toHaveLength(2);
    expect(ser.dataPoints?.[0]?.index).toBe(0);
    expect(ser.dataPoints?.[0]?.explosion).toBe(30);
    expect(ser.dataPoints?.[1]?.invertIfNegative).toBe(false);
  });

  it("round-trips a data-point marker override", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [
        {
          name: "S",
          values: [1, 2],
          dataPoints: [{ index: 1, marker: { symbol: "star", size: 10 } }],
        },
      ],
    };
    const result = roundTrip(opts);
    const ser = result.series[0] as ChartSeriesData;
    expect(ser.dataPoints?.[0]?.marker?.symbol).toBe("star");
    expect(ser.dataPoints?.[0]?.marker?.size).toBe(10);
  });

  it("series with enhancements is byte-stable on round-trip", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [
        {
          name: "S",
          values: [1, 2],
          marker: { symbol: "diamond" },
          smooth: false,
          dataPoints: [{ index: 0, explosion: 5 }],
        },
      ],
    };
    const xml1 = stringify(chartSpaceDesc, opts, {} as WriteContext);
    const result = roundTrip(opts);
    const xml2 = stringify(chartSpaceDesc, result, {} as WriteContext);
    expect(xml2).toBe(xml1);
  });

  it("round-trips 3D walls/floor thickness", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      threeD: true,
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      floor: { thickness: 25 },
      sideWall: { thickness: 30 },
      backWall: { thickness: 35 },
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('<c:floor><c:thickness val="25"/></c:floor>');
    expect(xml).toContain('<c:sideWall><c:thickness val="30"/></c:sideWall>');
    expect(xml).toContain('<c:backWall><c:thickness val="35"/></c:backWall>');

    const result = roundTrip(opts);
    expect(result.floor?.thickness).toBe(25);
    expect(result.sideWall?.thickness).toBe(30);
    expect(result.backWall?.thickness).toBe(35);
  });

  it("round-trips plot-area manual layout", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      plotAreaLayout: {
        layoutTarget: "inner",
        xMode: "edge",
        yMode: "edge",
        wMode: "edge",
        hMode: "edge",
        x: 0.1,
        y: 0.2,
        w: 0.7,
        h: 0.6,
      },
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain("<c:layout><c:manualLayout>");
    expect(xml).toContain('c:layoutTarget val="inner"');
    expect(xml).toContain('c:x val="0.1"');
    expect(xml).toContain('c:h val="0.6"');

    const result = roundTrip(opts);
    expect(result.plotAreaLayout?.layoutTarget).toBe("inner");
    expect(result.plotAreaLayout?.xMode).toBe("edge");
    expect(result.plotAreaLayout?.x).toBeCloseTo(0.1);
    expect(result.plotAreaLayout?.h).toBeCloseTo(0.6);
  });

  it("3D walls and manual layout are byte-stable on round-trip", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      threeD: true,
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      floor: { thickness: 25 },
      backWall: { thickness: 35 },
      plotAreaLayout: { layoutTarget: "inner", x: 0.1, y: 0.2 },
    };
    const xml1 = stringify(chartSpaceDesc, opts, {} as WriteContext);
    const result = roundTrip(opts);
    const xml2 = stringify(chartSpaceDesc, result, {} as WriteContext);
    expect(xml2).toBe(xml1);
  });

  it("round-trips bar gap width and overlap", () => {
    const opts: ChartSpaceOptions = {
      type: "column",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      gapWidth: 75,
      overlap: -20,
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('c:gapWidth val="75"');
    expect(xml).toContain('c:overlap val="-20"');

    const result = roundTrip(opts);
    expect(result.gapWidth).toBe(75);
    expect(result.overlap).toBe(-20);
  });

  it("round-trips pie first-slice angle and doughnut hole size", () => {
    const opts: ChartSpaceOptions = {
      type: "doughnut",
      categories: ["A", "B", "C"],
      series: [{ name: "S", values: [1, 2, 3] }],
      firstSliceAngle: 90,
      holeSize: 60,
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('c:firstSliceAng val="90"');
    expect(xml).toContain('c:holeSize val="60"');

    const result = roundTrip(opts);
    expect(result.firstSliceAngle).toBe(90);
    expect(result.holeSize).toBe(60);
  });

  it("round-trips bubble scale and negative-bubble options", () => {
    const opts: ChartSpaceOptions = {
      type: "bubble",
      series: [{ name: "S", xValues: [1], yValues: [2], bubbleSize: [3] }],
      bubbleScale: 80,
      showNegativeBubbles: true,
      sizeRepresents: "w",
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('c:bubbleScale val="80"');
    expect(xml).toContain("<c:showNegBubbles/>");
    expect(xml).toContain('c:sizeRepresents val="w"');

    const result = roundTrip(opts);
    expect(result.bubbleScale).toBe(80);
    expect(result.showNegativeBubbles).toBe(true);
    expect(result.sizeRepresents).toBe("w");
  });

  it("round-trips CT_Chart tail display options", () => {
    const opts: ChartSpaceOptions = {
      type: "line",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      plotVisOnly: false,
      displayBlanksAs: "span",
      showDataLabelsOverMax: true,
    };
    const xml = stringify(chartSpaceDesc, opts, {} as WriteContext);
    expect(xml).toContain('<c:plotVisOnly val="0"/>');
    expect(xml).toContain('c:dispBlanksAs val="span"');
    expect(xml).toContain("<c:showDLblsOverMax/>");

    const result = roundTrip(opts);
    expect(result.plotVisOnly).toBe(false);
    expect(result.displayBlanksAs).toBe("span");
    expect(result.showDataLabelsOverMax).toBe(true);
  });

  it("round-trips surface wireframe", () => {
    const opts: ChartSpaceOptions = {
      type: "surface",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
      wireframe: true,
    };
    const result = roundTrip(opts);
    expect(result.wireframe).toBe(true);
    const xml = stringify(chartSpaceDesc, result, {} as WriteContext);
    expect(xml).toContain("<c:wireframe/>");
  });
});
