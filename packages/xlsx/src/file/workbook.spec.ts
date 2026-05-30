import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import type { SheetDefinition } from "./workbook";
import { WorkbookXml } from "./workbook";

describe("WorkbookXml", () => {
  it("single sheet produces workbook with one sheet element", () => {
    const sheets: SheetDefinition[] = [{ name: "Sheet1", sheetId: 1, rId: "rId1" }];
    const tree = new Formatter().format(new WorkbookXml(sheets));

    const workbook = tree["workbook"] as any[];
    // _attr + bookViews + sheets + calcPr
    expect(workbook).toHaveLength(4);

    // Check namespace
    expect(workbook[0]["_attr"]["xmlns"]).toBe(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    );

    // Check sheet element (index 2: _attr=0, bookViews=1, sheets=2, calcPr=3)
    const sheetElements = workbook[2]["sheets"] as any[];
    expect(sheetElements).toHaveLength(1);
    expect(sheetElements[0]).to.deep.equal({
      sheet: { _attr: { name: "Sheet1", sheetId: "1", "r:id": "rId1" } },
    });
  });

  it("multiple sheets produce multiple sheet elements", () => {
    const sheets: SheetDefinition[] = [
      { name: "Sheet1", sheetId: 1, rId: "rId1" },
      { name: "Sheet2", sheetId: 2, rId: "rId2" },
      { name: "Sheet3", sheetId: 3, rId: "rId3" },
    ];
    const tree = new Formatter().format(new WorkbookXml(sheets));

    const workbook = tree["workbook"] as any[];
    const sheetElements = workbook[2]["sheets"] as any[];
    expect(sheetElements).toHaveLength(3);
  });

  it("hidden sheet includes state attribute", () => {
    const sheets: SheetDefinition[] = [
      { name: "Sheet1", sheetId: 1, rId: "rId1", state: "hidden" },
    ];
    const tree = new Formatter().format(new WorkbookXml(sheets));

    const workbook = tree["workbook"] as any[];
    const sheetElements = workbook[2]["sheets"] as any[];
    expect(sheetElements[0]["sheet"]["_attr"]["state"]).toBe("hidden");
  });

  it("visible sheet omits state attribute", () => {
    const sheets: SheetDefinition[] = [
      { name: "Sheet1", sheetId: 1, rId: "rId1", state: "visible" },
    ];
    const tree = new Formatter().format(new WorkbookXml(sheets));

    const workbook = tree["workbook"] as any[];
    const sheetElements = workbook[2]["sheets"] as any[];
    expect(sheetElements[0]["sheet"]["_attr"]["state"]).toBeUndefined();
  });
});
