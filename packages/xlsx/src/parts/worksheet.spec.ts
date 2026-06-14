import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { buildWorksheetXml } from "./worksheet";
import { worksheetDesc } from "./worksheet";
import type { WorksheetOptions } from "./worksheet";

describe("Worksheet", () => {
  describe("sheetProtection", () => {
    it("omits sheetProtection when not configured", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("sheetProtection");
    });

    it("generates sheetProtection with password hash", () => {
      const xml = buildWorksheetXml(
        {
          protection: { sheet: true, password: "test" },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain("<sheetProtection");
      expect(xml).toContain('sheet="1"');
      expect(xml).toContain("password=");
    });

    it("does not emit permission flags that match defaults", () => {
      const xml = buildWorksheetXml(
        {
          protection: { sheet: true },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      // formatCells defaults to true — only emitted when false
      expect(xml).not.toContain("formatCells");
      // selectLockedCells defaults to false — only emitted when true
      expect(xml).not.toContain("selectLockedCells");
    });

    it("emits formatCells=0 when explicitly set to false", () => {
      const xml = buildWorksheetXml(
        {
          protection: { sheet: true, formatCells: false },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('formatCells="0"');
    });

    it("emits selectLockedCells=1 when set to true", () => {
      const xml = buildWorksheetXml(
        {
          protection: { sheet: true, selectLockedCells: true },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('selectLockedCells="1"');
    });

    it("handles multiple protection flags", () => {
      const xml = buildWorksheetXml(
        {
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
        },
        {},
      );
      expect(xml).toContain('sheet="1"');
      expect(xml).toContain('formatCells="0"');
      expect(xml).toContain('insertRows="0"');
      expect(xml).toContain('deleteColumns="0"');
      expect(xml).toContain('sort="0"');
      expect(xml).toContain('selectLockedCells="1"');
      expect(xml).toContain('selectUnlockedCells="1"');
    });

    it("produces self-closing element", () => {
      const xml = buildWorksheetXml(
        {
          protection: { sheet: true },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toMatch(/<sheetProtection[^/]*\/>/);
    });
  });

  describe("hyperlinks", () => {
    it("omits hyperlinks when not configured", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("hyperlinks");
    });

    it("generates external hyperlink with r:id", () => {
      const xml = buildWorksheetXml(
        {
          hyperlinks: [{ cell: "A1", target: { type: "external", url: "https://example.com" } }],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain("<hyperlinks>");
      expect(xml).toContain('ref="A1"');
      expect(xml).toContain('r:id="rId1"');
      expect(xml).toContain("</hyperlinks>");
    });

    it("generates internal hyperlink with location (no r:id)", () => {
      const xml = buildWorksheetXml(
        {
          hyperlinks: [{ cell: "B2", target: { type: "internal", location: "'Sheet2'!A1" } }],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('ref="B2"');
      expect(xml).toContain('location="&apos;Sheet2&apos;!A1"');
      expect(xml).not.toContain("r:id");
    });

    it("includes tooltip and display when provided", () => {
      const xml = buildWorksheetXml(
        {
          hyperlinks: [
            {
              cell: "A1",
              target: { type: "external", url: "https://example.com" },
              tooltip: "Visit Example",
              display: "Example Site",
            },
          ],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('tooltip="Visit Example"');
      expect(xml).toContain('display="Example Site"');
    });

    it("handles mixed external and internal hyperlinks", () => {
      const xml = buildWorksheetXml(
        {
          hyperlinks: [
            { cell: "A1", target: { type: "external", url: "https://a.com" } },
            { cell: "B1", target: { type: "internal", location: "Sheet2!A1" } },
            { cell: "C1", target: { type: "external", url: "https://b.com" } },
          ],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      // External links numbered 1,2 (internal skipped)
      expect(xml).toContain('r:id="rId1"');
      expect(xml).toContain('r:id="rId2"');
      // Internal link has no r:id
      expect(xml).toContain('location="Sheet2!A1"');
    });
  });

  describe("tabColor", () => {
    it("omits sheetPr when no tabColor", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("sheetPr");
    });

    it("generates sheetPr with tabColor rgb", () => {
      const xml = buildWorksheetXml(
        {
          tabColor: { rgb: "FF0000" },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<sheetPr><tabColor rgb="FF0000"/></sheetPr>');
    });

    it("supports theme and tint", () => {
      const xml = buildWorksheetXml(
        {
          tabColor: { theme: 2, tint: 0.5 },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('theme="2"');
      expect(xml).toContain('tint="0.5"');
    });
  });

  describe("pageSetup", () => {
    it("omits pageSetup when not configured", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("pageSetup");
    });

    it("generates pageSetup with orientation", () => {
      const xml = buildWorksheetXml(
        {
          pageSetup: { orientation: "landscape" },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('orientation="landscape"');
    });

    it("generates pageSetup with paperSize", () => {
      const xml = buildWorksheetXml(
        {
          pageSetup: { paperSize: 9 },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('paperSize="9"');
    });

    it("generates pageSetup with fitToWidth/Height", () => {
      const xml = buildWorksheetXml(
        {
          pageSetup: { fitToWidth: 1, fitToHeight: 0 },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('fitToWidth="1"');
      expect(xml).toContain('fitToHeight="0"');
    });
  });

  describe("headerFooter", () => {
    it("omits headerFooter when not configured", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("headerFooter");
    });

    it("generates headerFooter with oddHeader/oddFooter", () => {
      const xml = buildWorksheetXml(
        {
          headerFooter: { oddHeader: "Page &P", oddFooter: "Confidential" },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain("<oddHeader>Page &amp;P</oddHeader>");
      expect(xml).toContain("<oddFooter>Confidential</oddFooter>");
    });

    it("supports differentOddEven flag", () => {
      const xml = buildWorksheetXml(
        {
          headerFooter: { oddHeader: "Odd", evenHeader: "Even", differentOddEven: true },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('differentOddEven="1"');
      expect(xml).toContain("<evenHeader>Even</evenHeader>");
    });
  });

  describe("sheetView", () => {
    it("defaults to tabSelected=1", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).toContain('tabSelected="1"');
    });

    it("omits default attributes", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("showGridLines");
      expect(xml).not.toContain("showZeros");
    });

    it("supports showGridLines=false", () => {
      const xml = buildWorksheetXml(
        {
          sheetView: { showGridLines: false },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('showGridLines="0"');
    });

    it("supports zoomScale", () => {
      const xml = buildWorksheetXml(
        {
          sheetView: { zoomScale: 150 },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('zoomScale="150"');
    });

    it("supports rightToLeft", () => {
      const xml = buildWorksheetXml(
        {
          sheetView: { rightToLeft: true },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('rightToLeft="1"');
    });
  });

  describe("column grouping", () => {
    it("includes outlineLevel in col element", () => {
      const xml = buildWorksheetXml(
        {
          columns: [{ min: 2, max: 3, outlineLevel: 1, hidden: true }],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('outlineLevel="1"');
      expect(xml).toContain('hidden="1"');
    });

    it("includes collapsed in col element", () => {
      const xml = buildWorksheetXml(
        {
          columns: [{ min: 1, max: 1, collapsed: true }],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('collapsed="1"');
    });

    it("outputs outlinePr when columns have outlineLevel", () => {
      const xml = buildWorksheetXml(
        {
          columns: [{ min: 2, max: 3, outlineLevel: 1 }],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<outlinePr summaryBelow="1" summaryRight="1"/>');
    });
  });

  describe("advanced autoFilter", () => {
    it("still supports string shorthand", () => {
      const xml = buildWorksheetXml(
        { autoFilter: "A1:D10", rows: [{ cells: [{ value: "A" }] }] },
        {},
      );
      expect(xml).toContain('<autoFilter ref="A1:D10"/>');
    });

    it("generates top10 filter", () => {
      const xml = buildWorksheetXml(
        {
          autoFilter: { ref: "A1:D10", top10: [{ colId: 2, val: 5 }] },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<filterColumn colId="2">');
      expect(xml).toContain('<top10 val="5"/>');
    });

    it("generates customFilters", () => {
      const xml = buildWorksheetXml(
        {
          autoFilter: {
            ref: "A1:D10",
            customFilters: [{ colId: 1, operator: "greaterThan", val: "100" }],
          },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<customFilter val="100" operator="greaterThan"/>');
    });

    it("generates sortState", () => {
      const xml = buildWorksheetXml(
        {
          autoFilter: { ref: "A1:D10", sort: [{ ref: "B1", descending: true }] },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<sortState ref="A1:D10">');
      expect(xml).toContain('ref="B1"');
      expect(xml).toContain('descending="1"');
    });
  });

  describe("sheetView round-trip", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
      sharedStrings: [],
    } as unknown as ReadContext;

    function roundTrip(opts: WorksheetOptions) {
      const xml = buildWorksheetXml(opts, {});
      const doc = parseXml(xml);
      const el = doc.elements![0];
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("round-trips sheetView display + zoom fields", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        sheetView: {
          windowProtection: true,
          showFormulas: true,
          showRuler: false,
          showOutlineSymbols: false,
          defaultGridColor: false,
          showWhiteSpace: false,
          colorId: 5,
          zoomScaleNormal: 80,
          zoomScaleSheetLayoutView: 70,
          zoomScalePageLayoutView: 60,
        },
      };
      const result = roundTrip(opts);
      const sv = result.sheetView!;

      expect(sv.windowProtection).toBe(true);
      expect(sv.showFormulas).toBe(true);
      expect(sv.showRuler).toBe(false);
      expect(sv.showOutlineSymbols).toBe(false);
      expect(sv.defaultGridColor).toBe(false);
      expect(sv.showWhiteSpace).toBe(false);
      expect(sv.colorId).toBe(5);
      expect(sv.zoomScaleNormal).toBe(80);
      expect(sv.zoomScaleSheetLayoutView).toBe(70);
      expect(sv.zoomScalePageLayoutView).toBe(60);
    });
  });

  describe("sheetPr round-trip", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
      sharedStrings: [],
    } as unknown as ReadContext;

    function roundTrip(opts: WorksheetOptions) {
      const xml = buildWorksheetXml(opts, {});
      const doc = parseXml(xml);
      const el = doc.elements![0];
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("round-trips outlinePr summaryBelow/summaryRight", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        columns: [{ min: 1, max: 1, outlineLevel: 1 }],
        sheetPr: { outlineSummaryBelow: false, outlineSummaryRight: false },
      };
      const result = roundTrip(opts);
      const sp = result.sheetPr!;

      expect(sp.outlineSummaryBelow).toBe(false);
      expect(sp.outlineSummaryRight).toBe(false);
    });

    it("round-trips pageSetUpPr fitToPage", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        pageSetup: { fitToWidth: 1, fitToHeight: 1 },
      };
      const result = roundTrip(opts);

      expect(result.pageSetup?.fitToPage).toBe(true);
    });
  });
});
