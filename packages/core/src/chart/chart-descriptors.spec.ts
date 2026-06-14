import { parse as parseXml } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { stringify, parse } from "../descriptor";
import { chartSpaceDesc } from "./chart-descriptors";
import type { ChartSpaceOptions } from "./types";

function roundTrip(opts: ChartSpaceOptions): ChartSpaceOptions {
  const xml = stringify(chartSpaceDesc, opts, {} as any);
  if (!xml) throw new Error("stringify returned undefined");
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return parse(chartSpaceDesc, el, {} as any);
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
    const xml = stringify(chartSpaceDesc, opts, {} as any);
    expect(xml).toContain("c:chartSpace");
    expect(xml).toContain("c:chart");
    expect(xml).toContain("c:plotArea");
    expect(xml).toContain("c:barChart");
    expect(xml).toContain("c:ser");
    expect(xml).toContain("Sales");
  });
});
