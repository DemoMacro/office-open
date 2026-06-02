import { describe, expect, it } from "vite-plus/test";

import { Worksheet } from "./worksheet";

const context = { stack: [] };

describe("Worksheet", () => {
  describe("sheetProtection", () => {
    it("omits sheetProtection when not configured", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).not.toContain("sheetProtection");
    });

    it("generates sheetProtection with password hash", () => {
      const ws = new Worksheet({
        protection: { sheet: true, password: "test" },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain("<sheetProtection");
      expect(xml).toContain('sheet="1"');
      expect(xml).toContain("password=");
    });

    it("does not emit permission flags that match defaults", () => {
      const ws = new Worksheet({
        protection: { sheet: true },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      // formatCells defaults to true — only emitted when false
      expect(xml).not.toContain("formatCells");
      // selectLockedCells defaults to false — only emitted when true
      expect(xml).not.toContain("selectLockedCells");
    });

    it("emits formatCells=0 when explicitly set to false", () => {
      const ws = new Worksheet({
        protection: { sheet: true, formatCells: false },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('formatCells="0"');
    });

    it("emits selectLockedCells=1 when set to true", () => {
      const ws = new Worksheet({
        protection: { sheet: true, selectLockedCells: true },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('selectLockedCells="1"');
    });

    it("handles multiple protection flags", () => {
      const ws = new Worksheet({
        protection: {
          sheet: true,
          password: "secret",
          formatCells: false,
          insertRows: false,
          deleteColumns: false,
          sort: false,
          selectLockedCells: true,
          selectUnlockedCells: true,
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('sheet="1"');
      expect(xml).toContain('formatCells="0"');
      expect(xml).toContain('insertRows="0"');
      expect(xml).toContain('deleteColumns="0"');
      expect(xml).toContain('sort="0"');
      expect(xml).toContain('selectLockedCells="1"');
      expect(xml).toContain('selectUnlockedCells="1"');
    });

    it("produces self-closing element", () => {
      const ws = new Worksheet({
        protection: { sheet: true },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toMatch(/<sheetProtection[^/]*\/>/);
    });
  });

  describe("hyperlinks", () => {
    it("omits hyperlinks when not configured", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).not.toContain("hyperlinks");
    });

    it("generates external hyperlink with r:id", () => {
      const ws = new Worksheet({
        hyperlinks: [{ cell: "A1", target: { type: "external", url: "https://example.com" } }],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain("<hyperlinks>");
      expect(xml).toContain('ref="A1"');
      expect(xml).toContain('r:id="rId1"');
      expect(xml).toContain("</hyperlinks>");
    });

    it("generates internal hyperlink with location (no r:id)", () => {
      const ws = new Worksheet({
        hyperlinks: [{ cell: "B2", target: { type: "internal", location: "'Sheet2'!A1" } }],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('ref="B2"');
      expect(xml).toContain('location="&apos;Sheet2&apos;!A1"');
      expect(xml).not.toContain("r:id");
    });

    it("includes tooltip and display when provided", () => {
      const ws = new Worksheet({
        hyperlinks: [
          {
            cell: "A1",
            target: { type: "external", url: "https://example.com" },
            tooltip: "Visit Example",
            display: "Example Site",
          },
        ],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('tooltip="Visit Example"');
      expect(xml).toContain('display="Example Site"');
    });

    it("handles mixed external and internal hyperlinks", () => {
      const ws = new Worksheet({
        hyperlinks: [
          { cell: "A1", target: { type: "external", url: "https://a.com" } },
          { cell: "B1", target: { type: "internal", location: "Sheet2!A1" } },
          { cell: "C1", target: { type: "external", url: "https://b.com" } },
        ],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      // External links numbered 1,2 (internal skipped)
      expect(xml).toContain('r:id="rId1"');
      expect(xml).toContain('r:id="rId2"');
      // Internal link has no r:id
      expect(xml).toContain('location="Sheet2!A1"');
    });
  });

  describe("tabColor", () => {
    it("omits sheetPr when no tabColor", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).not.toContain("sheetPr");
    });

    it("generates sheetPr with tabColor rgb", () => {
      const ws = new Worksheet({
        tabColor: { rgb: "FF0000" },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('<sheetPr><tabColor rgb="FF0000"/></sheetPr>');
    });

    it("supports theme and tint", () => {
      const ws = new Worksheet({
        tabColor: { theme: 2, tint: 0.5 },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('theme="2"');
      expect(xml).toContain('tint="0.5"');
    });
  });

  describe("pageSetup", () => {
    it("omits pageSetup when not configured", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).not.toContain("pageSetup");
    });

    it("generates pageSetup with orientation", () => {
      const ws = new Worksheet({
        pageSetup: { orientation: "landscape" },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('orientation="landscape"');
    });

    it("generates pageSetup with paperSize", () => {
      const ws = new Worksheet({
        pageSetup: { paperSize: 9 },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('paperSize="9"');
    });

    it("generates pageSetup with fitToWidth/Height", () => {
      const ws = new Worksheet({
        pageSetup: { fitToWidth: 1, fitToHeight: 0 },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('fitToWidth="1"');
      expect(xml).toContain('fitToHeight="0"');
    });
  });

  describe("headerFooter", () => {
    it("omits headerFooter when not configured", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).not.toContain("headerFooter");
    });

    it("generates headerFooter with oddHeader/oddFooter", () => {
      const ws = new Worksheet({
        headerFooter: { oddHeader: "Page &P", oddFooter: "Confidential" },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain("<oddHeader>Page &amp;P</oddHeader>");
      expect(xml).toContain("<oddFooter>Confidential</oddFooter>");
    });

    it("supports differentOddEven flag", () => {
      const ws = new Worksheet({
        headerFooter: { oddHeader: "Odd", evenHeader: "Even", differentOddEven: true },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('differentOddEven="1"');
      expect(xml).toContain("<evenHeader>Even</evenHeader>");
    });
  });

  describe("sheetView", () => {
    it("defaults to tabSelected=1", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).toContain('tabSelected="1"');
    });

    it("omits default attributes", () => {
      const ws = new Worksheet({ rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).not.toContain("showGridLines");
      expect(xml).not.toContain("showZeros");
    });

    it("supports showGridLines=false", () => {
      const ws = new Worksheet({
        sheetView: { showGridLines: false },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('showGridLines="0"');
    });

    it("supports zoomScale", () => {
      const ws = new Worksheet({
        sheetView: { zoomScale: 150 },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('zoomScale="150"');
    });

    it("supports rightToLeft", () => {
      const ws = new Worksheet({
        sheetView: { rightToLeft: true },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('rightToLeft="1"');
    });
  });

  describe("column grouping", () => {
    it("includes outlineLevel in col element", () => {
      const ws = new Worksheet({
        columns: [{ min: 2, max: 3, outlineLevel: 1, hidden: true }],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('outlineLevel="1"');
      expect(xml).toContain('hidden="1"');
    });

    it("includes collapsed in col element", () => {
      const ws = new Worksheet({
        columns: [{ min: 1, max: 1, collapsed: true }],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('collapsed="1"');
    });

    it("outputs outlinePr when columns have outlineLevel", () => {
      const ws = new Worksheet({
        columns: [{ min: 2, max: 3, outlineLevel: 1 }],
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('<outlinePr summaryBelow="1" summaryRight="1"/>');
    });
  });

  describe("advanced autoFilter", () => {
    it("still supports string shorthand", () => {
      const ws = new Worksheet({ autoFilter: "A1:D10", rows: [{ cells: [{ value: "A" }] }] });
      const xml = ws.toXml(context);
      expect(xml).toContain('<autoFilter ref="A1:D10"/>');
    });

    it("generates top10 filter", () => {
      const ws = new Worksheet({
        autoFilter: { ref: "A1:D10", top10: [{ colId: 2, val: 5 }] },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('<filterColumn colId="2">');
      expect(xml).toContain('<top10 val="5"/>');
    });

    it("generates customFilters", () => {
      const ws = new Worksheet({
        autoFilter: {
          ref: "A1:D10",
          customFilters: [{ colId: 1, operator: "greaterThan", val: "100" }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('<customFilter val="100" operator="greaterThan"/>');
    });

    it("generates sortState", () => {
      const ws = new Worksheet({
        autoFilter: { ref: "A1:D10", sort: [{ ref: "B1", descending: true }] },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const xml = ws.toXml(context);
      expect(xml).toContain('<sortState ref="A1:D10">');
      expect(xml).toContain('ref="B1"');
      expect(xml).toContain('descending="1"');
    });
  });
});
