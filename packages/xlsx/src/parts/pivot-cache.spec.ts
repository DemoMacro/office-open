import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { pivotCacheDefDesc, pivotCacheRecordsDesc } from "./pivot-cache";
import type {
  PivotCacheDefDescriptorOptions,
  PivotCacheRecordsDescriptorOptions,
} from "./pivot-cache";

// ── Minimal context stubs ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTripDef(opts: PivotCacheDefDescriptorOptions) {
  const xml = pivotCacheDefDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return pivotCacheDefDesc.parse(el, readCtx) as unknown as Record<string, unknown>;
}

function roundTripRecords(opts: PivotCacheRecordsDescriptorOptions) {
  const xml = pivotCacheRecordsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return pivotCacheRecordsDesc.parse(el, readCtx) as unknown as Record<string, unknown>;
}

// ── Tests ──

describe("pivotCacheDefDesc round-trip", () => {
  const baseSourceData = {
    fieldNames: ["Name", "Amount"],
    records: [
      ["Alice", 100],
      ["Bob", 200],
    ],
  };

  it("round-trips minimal pivot cache definition", () => {
    const opts: PivotCacheDefDescriptorOptions = {
      sourceRef: "A1:B3",
      sourceSheet: "Sheet1",
      sourceData: baseSourceData,
      recordsRid: "rId1",
    };
    const result = roundTripDef(opts);

    const ws = result.worksheetSource as Record<string, unknown>;
    expect(ws.ref).toBe("A1:B3");
    expect(ws.sheet).toBe("Sheet1");

    const fields = result.cacheFields as Record<string, unknown>[];
    expect(fields).toHaveLength(2);
    expect(fields[0].name).toBe("Name");
    expect(fields[1].name).toBe("Amount");
  });

  it("round-trips cache fields with shared items", () => {
    const opts: PivotCacheDefDescriptorOptions = {
      sourceRef: "A1:B3",
      sourceSheet: "Sheet1",
      sourceData: baseSourceData,
      recordsRid: "rId1",
    };
    const result = roundTripDef(opts);

    const fields = result.cacheFields as Record<string, unknown>[];
    // "Name" field has string shared items
    const nameItems = fields[0].sharedItems as (string | number)[];
    expect(nameItems).toContain("Alice");
    expect(nameItems).toContain("Bob");

    // "Amount" field has numeric shared items
    const amountItems = fields[1].sharedItems as (string | number)[];
    expect(amountItems).toContain(100);
    expect(amountItems).toContain(200);
  });
});

describe("pivotCacheRecordsDesc round-trip", () => {
  it("round-trips records with string and numeric fields", () => {
    const sourceData = {
      fieldNames: ["Name", "Amount"],
      records: [
        ["Alice", 100],
        ["Bob", 200],
      ],
    };
    const opts: PivotCacheRecordsDescriptorOptions = { sourceData };
    const result = roundTripRecords(opts);

    const records = result.records as Record<string, unknown>[][];
    expect(records).toHaveLength(2);

    // First record: string index + number
    expect(records[0]).toHaveLength(2);
    expect(records[0][0].type).toBe("string");
    expect(records[0][1].type).toBe("number");
    expect(records[0][1].v).toBe(100);

    // Second record
    expect(records[1]).toHaveLength(2);
    expect(records[1][1].v).toBe(200);
  });
});
