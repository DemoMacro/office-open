import { describe, expect, it } from "vite-plus/test";

import {
  type ChartDocumentOptions,
  generateChartDocument,
  generateDatabaseDocument,
  parseChartDocument,
  parseDatabaseDocument,
} from "./index";

describe("ODF schema codecs", () => {
  it("round-trips chart subdocuments through the shared chart model", () => {
    const chart: ChartDocumentOptions = {
      title: "Chart",
      chart: { type: "column", series: [{ name: "Sales", values: [1, 2] }] },
    };
    const parsed = parseChartDocument(generateChartDocument(chart));
    expect(parsed.title).toBe(chart.title);
    expect(parsed.chart.type).toBe("column");
    expect(parsed.chart.series[0]).toMatchObject(chart.chart.series[0]!);
  });

  it("round-trips database subdocuments as generic ODF nodes", () => {
    const database = {
      title: "Database",
      body: [{ name: "db:data-source", attributes: { "db:name": "Library" } }],
    };
    expect(parseDatabaseDocument(generateDatabaseDocument(database))).toEqual(database);
  });
});
