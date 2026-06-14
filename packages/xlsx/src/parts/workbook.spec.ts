import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { workbookDesc } from "./workbook";
import type { WorkbookDescriptorOptions } from "./workbook";

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

function roundTrip(opts: WorkbookDescriptorOptions) {
  const xml = workbookDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return workbookDesc.parse(el, readCtx) as unknown as WorkbookDescriptorOptions;
}

// ── Tests ──

describe("workbookDesc round-trip", () => {
  it("round-trips minimal workbook with sheets", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [
        { name: "Sheet1", sheetId: 1, rId: "rId1" },
        { name: "Sheet2", sheetId: 2, rId: "rId2" },
      ],
    };
    const result = roundTrip(opts);

    expect(result.sheets).toHaveLength(2);
    expect(result.sheets[0].name).toBe("Sheet1");
    expect(result.sheets[0].sheetId).toBe(1);
    expect(result.sheets[0].rId).toBe("rId1");
    expect(result.sheets[1].name).toBe("Sheet2");
  });

  it("round-trips sheet with hidden state", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Hidden", sheetId: 1, rId: "rId1", state: "hidden" }],
    };
    const result = roundTrip(opts);

    expect(result.sheets[0].state).toBe("hidden");
  });

  it("round-trips workbook protection", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      protection: { lockStructure: true, lockWindows: true },
    };
    const result = roundTrip(opts);

    expect(result.protection?.lockStructure).toBe(true);
    expect(result.protection?.lockWindows).toBe(true);
  });

  it("round-trips book view", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      bookView: {
        activeTab: 1,
        xWindow: 100,
        yWindow: 200,
        windowWidth: 28800,
        windowHeight: 12300,
        autoFilterDateGrouping: false,
        firstSheet: 3,
        tabRatio: 400,
      },
    };
    const result = roundTrip(opts);

    expect(result.bookView?.activeTab).toBe(1);
    expect(result.bookView?.xWindow).toBe(100);
    expect(result.bookView?.yWindow).toBe(200);
    expect(result.bookView?.autoFilterDateGrouping).toBe(false);
    expect(result.bookView?.firstSheet).toBe(3);
    expect(result.bookView?.tabRatio).toBe(400);
  });

  it("round-trips calc properties", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      calcPr: { calcId: 162913, calcMode: "auto", fullCalcOnLoad: true, refMode: "A1" },
    };
    const result = roundTrip(opts);

    expect(result.calcPr?.calcId).toBe(162913);
    expect(result.calcPr?.calcMode).toBe("auto");
    expect(result.calcPr?.fullCalcOnLoad).toBe(true);
    expect(result.calcPr?.refMode).toBe("A1");
  });

  it("round-trips workbookPr with date1904", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      workbookPr: { date1904: true },
    };
    const result = roundTrip(opts);

    expect(result.workbookPr?.date1904).toBe(true);
  });

  it("round-trips pivot caches", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      pivotCaches: [
        { cacheId: 1, rId: "rId5" },
        { cacheId: 2, rId: "rId6" },
      ],
    };
    const result = roundTrip(opts);

    expect(result.pivotCaches).toHaveLength(2);
    expect(result.pivotCaches![0].cacheId).toBe(1);
    expect(result.pivotCaches![1].rId).toBe("rId6");
  });

  it("round-trips function groups", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      functionGroups: ["UDF1", "UDF2"],
    };
    const result = roundTrip(opts);

    expect(result.functionGroups).toEqual(["UDF1", "UDF2"]);
  });

  it("round-trips file sharing", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      fileSharing: { readOnlyRecommended: true, userName: "TestUser" },
    };
    const result = roundTrip(opts);

    expect(result.fileSharing?.readOnlyRecommended).toBe(true);
    expect(result.fileSharing?.userName).toBe("TestUser");
  });

  it("round-trips file recovery properties", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      fileRecoveryPr: { autoRecover: false, crashSave: true },
    };
    const result = roundTrip(opts);

    expect(result.fileRecoveryPr?.autoRecover).toBe(false);
    expect(result.fileRecoveryPr?.crashSave).toBe(true);
  });
});
