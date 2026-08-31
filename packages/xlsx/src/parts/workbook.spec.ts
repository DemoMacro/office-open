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
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
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
    expect(result.sheets[0]?.name).toBe("Sheet1");
    expect(result.sheets[0]?.sheetId).toBe(1);
    expect(result.sheets[0]?.rId).toBe("rId1");
    expect(result.sheets[1]?.name).toBe("Sheet2");
  });

  it("round-trips fileVersion", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      fileVersion: { appName: "xl", lastEdited: 8, lowestEdited: 7, rupBuild: 12345 },
    };
    const result = roundTrip(opts);

    expect(result.fileVersion).toEqual({
      appName: "xl",
      lastEdited: 8,
      lowestEdited: 7,
      rupBuild: 12345,
    });
  });

  it("maps customViews showComments to ST_Comments tokens", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      customViews: [
        {
          name: "Mine",
          guid: "{11111111-2222-3333-4444-555555555555}",
          windowWidth: 28800,
          windowHeight: 14400,
          activeSheetId: 1,
          showComments: "comment",
        },
      ],
    };
    const xml = workbookDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain('showComments="commIndAndComment"');

    const result = roundTrip(opts);
    expect(result.customViews?.[0]?.showComments).toBe("comment");
  });

  it("round-trips smartTagPr and smartTagTypes", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      smartTag: { embed: true, show: "noIndicator" },
      smartTagTypes: [
        {
          namespaceUri: "http://schemas.example.com/addr",
          name: "Address",
          url: "http://example.com",
        },
        { namespaceUri: "http://schemas.example.com/date", name: "Date" },
      ],
    };
    const result = roundTrip(opts);

    expect(result.smartTag).toEqual({ embed: true, show: "noIndicator" });
    expect(result.smartTagTypes).toEqual([
      {
        namespaceUri: "http://schemas.example.com/addr",
        name: "Address",
        url: "http://example.com",
      },
      { namespaceUri: "http://schemas.example.com/date", name: "Date" },
    ]);
  });

  it("omits smartTagPr/smartTagTypes when absent (defaults not emitted)", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
    };
    const result = roundTrip(opts);
    expect(result.smartTag).toBeUndefined();
    expect(result.smartTagTypes).toBeUndefined();
  });

  it("round-trips sheet with hidden state", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Hidden", sheetId: 1, rId: "rId1", state: "hidden" }],
    };
    const result = roundTrip(opts);

    expect(result.sheets[0]?.state).toBe("hidden");
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
      calculation: { calcId: 162913, calcMode: "auto", fullCalcOnLoad: true, refMode: "A1" },
    };
    const result = roundTrip(opts);

    expect(result.calculation?.calcId).toBe(162913);
    expect(result.calculation?.calcMode).toBe("auto");
    expect(result.calculation?.fullCalcOnLoad).toBe(true);
    expect(result.calculation?.refMode).toBe("A1");
  });

  it("round-trips workbookPr with date1904", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      properties: { date1904: true },
    };
    const result = roundTrip(opts);

    expect(result.properties?.date1904).toBe(true);
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
    expect(result.pivotCaches![0]?.cacheId).toBe(1);
    expect(result.pivotCaches![1]?.rId).toBe("rId6");
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
      fileRecovery: { autoRecover: false, crashSave: true },
    };
    const result = roundTrip(opts);

    expect(result.fileRecovery?.autoRecover).toBe(false);
    expect(result.fileRecovery?.crashSave).toBe(true);
  });

  it("round-trips defined names with full attribute set", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      definedNames: [
        {
          name: "TaxRate",
          value: "0.2",
          comment: "VAT rate",
          localSheetId: 0,
          hidden: true,
          functionGroupId: 1,
        },
        {
          name: "Total",
          value: "SUM(Sheet1!A1:A10)",
          customMenu: "Run Total",
          description: "Sum of sales",
          shortcutKey: "t",
          publishToServer: true,
        },
      ],
    };
    const result = roundTrip(opts);
    const dns = result.definedNames!;
    expect(dns).toHaveLength(2);
    expect(dns[0]?.name).toBe("TaxRate");
    expect(dns[0]?.value).toBe("0.2");
    expect(dns[0]?.comment).toBe("VAT rate");
    expect(dns[0]?.localSheetId).toBe(0);
    expect(dns[0]?.hidden).toBe(true);
    expect(dns[0]?.functionGroupId).toBe(1);
    expect(dns[1]?.name).toBe("Total");
    expect(dns[1]?.value).toBe("SUM(Sheet1!A1:A10)");
    expect(dns[1]?.customMenu).toBe("Run Total");
    expect(dns[1]?.publishToServer).toBe(true);
  });

  it("emits definedNames after externalReferences and before calcPr (XSD sequence)", () => {
    const xml = workbookDesc.stringify(
      {
        sheets: [{ name: "S", sheetId: 1, rId: "rId1" }],
        definedNames: [{ name: "X", value: "1" }],
        calculation: { calcId: 1 },
      },
      writeCtx,
    )!;
    expect(xml.indexOf("<definedNames>")).toBeLessThan(xml.indexOf("<calcPr"));
    expect(xml.indexOf("EXTERNAL_REFS")).toBeLessThan(xml.indexOf("<definedNames>"));
  });

  it("round-trips coauthoring revision state", () => {
    const opts: WorkbookDescriptorOptions = {
      sheets: [{ name: "Sheet1", sheetId: 1, rId: "rId1" }],
      revisionPtr: {
        revisionIdLastSave: 0,
        documentId: "8_{5FF0C957-174C-468D-A376-EA8B81D2939C}",
        coauthVersionLast: 47,
        coauthVersionMax: 47,
        uidLastSave: "{00000000-0000-0000-0000-000000000000}",
      },
    };
    const xml = workbookDesc.stringify(opts, writeCtx)!;
    expect(xml).toMatch(/<xr:revisionPtr revIDLastSave="0"/);
    expect(xml).toMatch(/xr6:coauthVersionMax="47"/);

    const result = roundTrip(opts);
    expect(result.revisionPtr).toEqual(opts.revisionPtr);
  });

  it("parses the x15 absPath form alongside x15ac", () => {
    const xml =
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
      ' xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"' +
      ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15">' +
      '<mc:AlternateContent><mc:Choice Requires="x15">' +
      '<x15:absPath url="C:\\Users\\kazuma\\Desktop\\"/></mc:Choice></mc:AlternateContent>' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("no root element");
    const result = workbookDesc.parse(el, readCtx) as unknown as WorkbookDescriptorOptions;
    expect(result.absPath).toBe("C:\\Users\\kazuma\\Desktop\\");
  });
});
