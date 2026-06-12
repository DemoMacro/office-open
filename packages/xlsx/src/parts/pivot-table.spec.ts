import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { pivotTableDesc } from "./pivot-table";
import type { PivotTableDescriptorOptions } from "./pivot-table";
import type { PivotSourceData } from "./pivot/pivot-utils";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

interface PivotTableParseResult extends PivotTableDescriptorOptions {
  name?: string;
  location?: string;
  pivotFields?: Array<{ axis?: string }>;
  dataFields?: Array<{ name?: string; subtotal?: string }>;
  style?: string;
}

function roundTrip(opts: PivotTableDescriptorOptions) {
  const xml = pivotTableDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return pivotTableDesc.parse(el, readCtx) as unknown as PivotTableParseResult;
}

// ── Shared source data ──

const sourceData: PivotSourceData = {
  fieldNames: ["Region", "Product", "Sales"],
  records: [
    ["East", "A", 100],
    ["East", "B", 200],
    ["West", "A", 150],
    ["West", "B", 250],
  ],
};

// ── Tests ──

describe("pivotTableDesc round-trip", () => {
  it("round-trips basic pivot with rows and data", () => {
    const opts: PivotTableDescriptorOptions = {
      options: {
        source: "Sheet1!A1:C5",
        rows: ["Region"],
        data: [{ field: "Sales", summarize: "sum" }],
      },
      sourceData,
      cacheId: 1,
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("PivotTable1");
    expect(result.cacheId).toBe(1);
    expect(result.location).toBeDefined();
    expect(typeof result.location).toBe("string");

    const pivotFields = result.pivotFields!;
    expect(pivotFields).toHaveLength(3);
    // Region is axisRow
    expect(pivotFields[0].axis).toBe("axisRow");
    // Sales is dataField
    expect(pivotFields[2].axis).toBeUndefined();
  });

  it("round-trips pivot with custom name", () => {
    const opts: PivotTableDescriptorOptions = {
      options: {
        name: "MyPivot",
        source: "A1:C5",
        rows: ["Region"],
        data: [{ field: "Sales" }],
      },
      sourceData,
      cacheId: 5,
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("MyPivot");
    expect(result.cacheId).toBe(5);
  });

  it("round-trips data fields with subtotal type", () => {
    const opts: PivotTableDescriptorOptions = {
      options: {
        source: "A1:C5",
        rows: ["Region"],
        data: [
          { field: "Sales", summarize: "average", name: "Avg Sales" },
          { field: "Sales", summarize: "count", name: "Count" },
        ],
      },
      sourceData,
      cacheId: 1,
    };
    const result = roundTrip(opts);

    const dataFields = result.dataFields!;
    expect(dataFields).toHaveLength(2);
    expect(dataFields[0].name).toBe("Avg Sales");
    expect(dataFields[0].subtotal).toBe("average");
    expect(dataFields[1].name).toBe("Count");
    expect(dataFields[1].subtotal).toBe("count");
  });

  it("round-trips pivot with columns", () => {
    const opts: PivotTableDescriptorOptions = {
      options: {
        source: "A1:C5",
        rows: ["Region"],
        columns: ["Product"],
        data: [{ field: "Sales" }],
      },
      sourceData,
      cacheId: 1,
    };
    const result = roundTrip(opts);

    const pivotFields = result.pivotFields!;
    expect(pivotFields[1].axis).toBe("axisCol");
  });

  it("round-trips pivot with style", () => {
    const opts: PivotTableDescriptorOptions = {
      options: {
        source: "A1:C5",
        rows: ["Region"],
        data: [{ field: "Sales" }],
        style: "PivotStyleDark1",
      },
      sourceData,
      cacheId: 1,
    };
    const result = roundTrip(opts);

    expect(result.style).toBe("PivotStyleDark1");
  });

  it("round-trips pivot with multiple rows", () => {
    const opts: PivotTableDescriptorOptions = {
      options: {
        source: "A1:C5",
        rows: ["Region", "Product"],
        data: [{ field: "Sales" }],
      },
      sourceData,
      cacheId: 1,
    };
    const result = roundTrip(opts);

    const pivotFields = result.pivotFields!;
    expect(pivotFields[0].axis).toBe("axisRow");
    expect(pivotFields[1].axis).toBe("axisRow");
  });
});
