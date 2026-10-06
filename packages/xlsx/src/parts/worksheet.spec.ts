import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { SharedStrings } from "./shared-strings";
import { buildWorksheetXml } from "./worksheet";
import { worksheetDesc } from "./worksheet";
import type { AutoFilterOptions, WorksheetOptions } from "./worksheet";
import type { TabColorOptions } from "./worksheet/types";

describe("Worksheet", () => {
  describe("cell formula", () => {
    it("accepts a bare string as shorthand for { formula }", () => {
      const shorthand = buildWorksheetXml(
        { rows: [{ cells: [{ formula: "SUM(A1:A2)", reference: "B1" }] }] },
        {},
      );
      const expanded = buildWorksheetXml(
        { rows: [{ cells: [{ formula: { formula: "SUM(A1:A2)" }, reference: "B1" }] }] },
        {},
      );
      expect(shorthand).toContain("<f>SUM(A1:A2)</f>");
      expect(shorthand).toBe(expanded);
    });

    it("emits preserve for formulas with edge whitespace", () => {
      const xml = buildWorksheetXml(
        { rows: [{ cells: [{ formula: { formula: " SUM(A1) " }, reference: "B1" }] }] },
        {},
      );
      expect(xml).toContain('<f xml:space="preserve"> SUM(A1) </f>');
    });
  });

  describe("cell presence", () => {
    it("preserves a reference-only empty cell", () => {
      const xml = buildWorksheetXml(
        { rows: [{ rowNumber: 5, cells: [{ reference: "B5" }, { reference: "C5" }] }] },
        {},
      );
      expect(xml).toContain('<c r="B5"/>');
      expect(xml).toContain('<c r="C5"/>');
    });

    it("emits an empty shared-string cell without registering an item", () => {
      const sharedStrings = new SharedStrings();
      const xml = buildWorksheetXml(
        { rows: [{ cells: [{ reference: "A1", value: "" }] }] },
        { sharedStrings },
      );
      expect(xml).toContain('<c r="A1" t="s"><v/></c>');
      expect(sharedStrings.count).toBe(0);
    });

    it("omits an empty generated placeholder cell", () => {
      const xml = buildWorksheetXml({ rows: [{ cells: [{}] }] }, {});
      expect(xml).not.toContain("<c ");
    });
  });

  describe("sheetFormatPr presence", () => {
    it("preserves source absence", () => {
      const xml = buildWorksheetXml({ sheetFormat: false, rows: [] }, {});
      expect(xml).not.toContain("<sheetFormatPr");
    });

    it("emits fresh-authoring defaults only when unconfigured", () => {
      const xml = buildWorksheetXml({ rows: [] }, {});
      expect(xml).toContain('<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>');
    });
  });

  describe("worksheet source fidelity", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
      sharedStrings: [],
    } as unknown as ReadContext;

    function parseSource(xml: string): WorksheetOptions {
      const el = parseXml(xml, { nativeTypeAttributes: true }).elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("preserves explicit zero pane splits", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetViews><sheetView workbookViewId="0">` +
          `<pane xSplit="0" ySplit="0" topLeftCell="A1" activePane="topLeft" state="frozen"/>` +
          `</sheetView></sheetViews>` +
          `<sheetData/></worksheet>`,
      );
      expect(result.freezePanes).toEqual({
        row: 0,
        col: 0,
        state: "frozen",
        topLeftCell: "A1",
        activePane: "topLeft",
      });
      expect(buildWorksheetXml(result, {})).toContain(
        '<pane ySplit="0" xSplit="0" topLeftCell="A1" activePane="topLeft" state="frozen"/>',
      );
    });

    it("round-trips frozenSplit pane state", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetViews><sheetView workbookViewId="0">` +
          `<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozenSplit"/>` +
          `</sheetView></sheetViews>` +
          `<sheetData/></worksheet>`,
      );
      expect(result.freezePanes).toMatchObject({ state: "frozenSplit" });
      expect(buildWorksheetXml(result, {})).toContain('state="frozenSplit"');
    });

    it("keeps split shorthand compatible with explicit state", () => {
      const splitXml = buildWorksheetXml({ freezePanes: { row: 1, split: true } }, {});
      expect(splitXml).toContain('state="split"');
      const explicitXml = buildWorksheetXml(
        { freezePanes: { row: 1, split: true, state: "frozenSplit" } },
        {},
      );
      expect(explicitXml).toContain('state="frozenSplit"');
    });

    it("round-trips tabColor tint", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetPr><tabColor theme="2" tint="-0.25"/></sheetPr><sheetData/></worksheet>`,
      );
      expect(result.tabColor).toEqual({ theme: 2, tint: -0.25, tintRaw: "-0.25" });
      expect(buildWorksheetXml(result, {})).toContain('<tabColor theme="2" tint="-0.25"/>');
    });

    it("round-trips long tabColor tint lexemes", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetPr><tabColor theme="2" tint="-0.249977111117893"/></sheetPr><sheetData/></worksheet>`,
      );
      expect((result.tabColor as TabColorOptions).tintRaw).toBe("-0.249977111117893");
      expect(buildWorksheetXml(result, {})).toContain(
        '<tabColor theme="2" tint="-0.249977111117893"/>',
      );
    });

    it("preserves whitespace in validation and conditional formulas", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetData/><conditionalFormatting sqref="A1">` +
          `<cfRule type="cellIs" dxfId="0" priority="1" operator="between">` +
          `<formula> 1 </formula><formula> 10 </formula></cfRule></conditionalFormatting>` +
          `<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="B1">` +
          `<formula1> 1 </formula1><formula2> 10 </formula2></dataValidation></dataValidations>` +
          `</worksheet>`,
      );
      const xml = buildWorksheetXml(result, {});
      expect(xml).toContain('<formula xml:space="preserve"> 1 </formula>');
      expect(xml).toContain('<formula xml:space="preserve"> 10 </formula>');
      expect(xml).toContain('<formula1 xml:space="preserve"> 1 </formula1>');
      expect(xml).toContain('<formula2 xml:space="preserve"> 10 </formula2>');
    });

    it("round-trips revision uid and explicit filter mode", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
          `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" ` +
          `xr:uid="{A1000000-0000-0000-0000-000000000000}">` +
          `<sheetPr filterMode="0"/><sheetData/></worksheet>`,
      );
      expect(result.uid).toBe("{A1000000-0000-0000-0000-000000000000}");
      expect(result.properties?.filterMode).toBe(false);
      const xml = buildWorksheetXml(result, {});
      expect(xml).toContain('xr:uid="{A1000000-0000-0000-0000-000000000000}"');
      expect(xml).toContain('<sheetPr filterMode="0"/>');
    });

    it("round-trips sheetFormatPr and row dyDescent", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
          `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">` +
          `<sheetFormatPr x14ac:dyDescent="0.25"/>` +
          `<sheetData><row r="1" x14ac:dyDescent="0.25"><c r="A1" t="inlineStr">` +
          `<is><t>A</t></is></c></row></sheetData></worksheet>`,
      );
      const xml = buildWorksheetXml(result, {});
      expect(xml).toContain('x14ac:dyDescent="0.25"');
      expect(xml).toContain('<row r="1" x14ac:dyDescent="0.25"');
    });

    it("round-trips explicit sheetView defaults", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetViews><sheetView showGridLines="1" showRowColHeaders="0" showZeros="1" ` +
          `rightToLeft="0" windowProtection="0" showFormulas="0" showRuler="1" ` +
          `showOutlineSymbols="0" defaultGridColor="1" showWhiteSpace="0" workbookViewId="0"/>` +
          `</sheetViews><sheetData/></worksheet>`,
      );
      expect(result.sheetView).toMatchObject({
        showGridLines: true,
        showRowColHeaders: false,
        showZeros: true,
        rightToLeft: false,
        windowProtection: false,
        showFormulas: false,
        showRuler: true,
        showOutlineSymbols: false,
        defaultGridColor: true,
        showWhiteSpace: false,
      });
      const xml = buildWorksheetXml(result, {});
      expect(xml).toContain('showGridLines="1"');
      expect(xml).toContain('showRowColHeaders="0"');
      expect(xml).toContain('rightToLeft="0"');
      expect(xml).toContain('showWhiteSpace="0"');
    });

    it("keeps empty sheetPr and headerFooter elements", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<sheetPr/><sheetData/><headerFooter/></worksheet>`,
      );
      expect(result.properties).toEqual({});
      expect(result.headerFooter).toEqual({});
      const xml = buildWorksheetXml(result, {});
      expect(xml).toContain("<sheetPr/>");
      expect(xml).toContain("<headerFooter/>");
    });

    it("keeps a column source style index", () => {
      const result = parseSource(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<cols><col min="1" max="1" width="12" style="7"/></cols><sheetData/>` +
          `</worksheet>`,
      );
      expect(result.columns?.[0]?.style).toBe(7);
      expect(buildWorksheetXml(result, {})).toContain('style="7"');
    });
  });

  describe("conditional formatting formulas", () => {
    it("emits formulas before specialized rule content", () => {
      const formula = 'MAX(IF(A1="", 0, A1))';
      const xml = buildWorksheetXml(
        {
          rows: [{ cells: [{ value: 1 }] }],
          conditionalFormats: [
            {
              sqref: "A1:A10",
              rules: [
                {
                  type: "colorScale",
                  formulas: [formula],
                  colorScale: {
                    cfvo: [{ type: "min" }, { type: "max" }],
                    colors: [{ rgb: "FF0000" }, { rgb: "00FF00" }],
                  },
                },
                {
                  type: "dataBar",
                  formulas: [formula],
                  dataBar: {
                    cfvo: [{ type: "min" }, { type: "max" }],
                    color: { rgb: "638EC6" },
                  },
                },
              ],
            },
          ],
        },
        {},
      );
      const escapedFormula = "<formula>MAX(IF(A1=&quot;&quot;, 0, A1))</formula>";
      expect(xml).toContain(`${escapedFormula}<colorScale>`);
      expect(xml).toContain(`${escapedFormula}<dataBar>`);
    });

    it("round-trips a theme-channel data bar color without inventing rgb", () => {
      const readCtx = {
        resolveRelationship: () => undefined,
        getPart: () => undefined,
        getRaw: () => undefined,
      } as unknown as ReadContext;
      const el = parseXml(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<conditionalFormatting sqref="G1:G3"><cfRule type="dataBar" priority="5">` +
          `<dataBar><cfvo type="min" val="0"/><cfvo type="max" val="0"/>` +
          `<color theme="4"/></dataBar></cfRule></conditionalFormatting></worksheet>`,
      ).elements?.[0];
      if (!el) throw new Error("no root");
      const ws = worksheetDesc.parse(el, readCtx);
      expect(ws.conditionalFormats?.[0]?.rules[0]?.dataBar?.showValue).toBeUndefined();
      const xml = buildWorksheetXml(ws, {});
      expect(xml).toContain('<color theme="4"/>');
      expect(xml).not.toContain('rgb="FF"');
    });

    it("round-trips a hidden data bar value", () => {
      const readCtx = {
        resolveRelationship: () => undefined,
        getPart: () => undefined,
        getRaw: () => undefined,
      } as unknown as ReadContext;
      const el = parseXml(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<conditionalFormatting sqref="G1:G3"><cfRule type="dataBar" priority="5">` +
          `<dataBar showValue="0"><cfvo type="min" val="0"/><cfvo type="max" val="0"/>` +
          `<color theme="4"/></dataBar></cfRule></conditionalFormatting></worksheet>`,
      ).elements?.[0];
      if (!el) throw new Error("no root");
      const ws = worksheetDesc.parse(el, readCtx);
      const xml = buildWorksheetXml(ws, {});
      expect(ws.conditionalFormats?.[0]?.rules[0]?.dataBar?.showValue).toBe(false);
      expect(xml).toContain('<dataBar showValue="0">');
    });
  });

  it("preserves an empty formula cache as a string result", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const el = parseXml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<sheetData><row r="1"><c r="A1" t="str"><f>IFERROR(A2,"")</f><v/></c></row></sheetData>` +
        `</worksheet>`,
    ).elements?.[0];
    if (!el) throw new Error("no root");
    const ws = worksheetDesc.parse(el, readCtx);
    const cell = ws.rows?.[0]?.cells?.[0];
    expect(cell?.value).toBe("");
    expect(buildWorksheetXml(ws, {})).toContain(
      'r="A1" t="str"><f>IFERROR(A2,&quot;&quot;)</f><v></v>',
    );
  });

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
          hyperlinks: [{ cell: "A1", url: "https://example.com" }],
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
          hyperlinks: [{ cell: "B2", location: "'Sheet2'!A1" }],
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
              url: "https://example.com",
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

    it("preserves coauthoring identity", () => {
      const readContext = {
        resolveRelationship: () => undefined,
        getPart: () => undefined,
        getRaw: () => undefined,
      } as unknown as ReadContext;
      const el = parseXml(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
          ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
          ` xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
          `<hyperlinks><hyperlink ref="A1" location="Sheet2!A1" xr:uid="{00000000-0000-0000-0000-000000000000}"/>` +
          `</hyperlinks></worksheet>`,
      ).elements?.[0];
      if (!el) throw new Error("no root");
      const result = worksheetDesc.parse(el, readContext);
      expect(result.hyperlinks?.[0]?.uid).toBe("{00000000-0000-0000-0000-000000000000}");
      expect(buildWorksheetXml(result, {})).toContain(
        'xr:uid="{00000000-0000-0000-0000-000000000000}"',
      );
    });

    it("emits both r:id and location when url+location are set together", () => {
      // CT_Hyperlink's @r:id and @location are independent — an external
      // workbook plus an internal jump target is a legal combination.
      const xml = buildWorksheetXml(
        {
          hyperlinks: [{ cell: "A1", url: "https://example.com", location: "Sheet2!A1" }],
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('r:id="rId1"');
      expect(xml).toContain('location="Sheet2!A1"');
    });

    it("handles mixed external and internal hyperlinks", () => {
      const xml = buildWorksheetXml(
        {
          hyperlinks: [
            { cell: "A1", url: "https://a.com" },
            { cell: "B1", location: "Sheet2!A1" },
            { cell: "C1", url: "https://b.com" },
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

  it("preserves page-margin lexical precision", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const el = parseXml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<pageMargins left="0.70866141732283472" right="0.70866141732283472"` +
        ` top="0.74803149606299213" bottom="0.74803149606299213"` +
        ` header="0.31496062992125984" footer="0.31496062992125984"/>` +
        `</worksheet>`,
    ).elements?.[0];
    if (!el) throw new Error("no root");
    const ws = worksheetDesc.parse(el, readCtx);
    expect(ws.pageMargins?.leftRaw).toBe("0.70866141732283472");
    expect(buildWorksheetXml(ws, {})).toContain(
      '<pageMargins left="0.70866141732283472" right="0.70866141732283472"',
    );
  });

  it("round-trips data consolidation references and counts", () => {
    const readCtx = {
      resolveExternalImage: (rId: string) =>
        rId === "rId1" ? "/sources/consolidation.xlsm" : undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const el = parseXml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
        ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<dataConsolidate function="average" topLabels="1">` +
        `<dataRefs count="2">` +
        `<dataRef name="SourceArea" sheet="Sheet1"/>` +
        `<dataRef ref="A1:B2" sheet="average" r:id="rId1"/>` +
        `</dataRefs></dataConsolidate></worksheet>`,
    ).elements?.[0];
    if (!el) throw new Error("no root");
    const ws = worksheetDesc.parse(el, readCtx);
    expect(ws.dataConsolidate).toEqual({
      function: "average",
      topLabels: true,
      count: 2,
      refs: [
        { name: "SourceArea", sheet: "Sheet1" },
        { ref: "A1:B2", sheet: "average", rId: "rId1", target: "/sources/consolidation.xlsm" },
      ],
    });
    expect(buildWorksheetXml(ws, {})).toContain(
      '<dataRefs count="2"><dataRef sheet="Sheet1" name="SourceArea"/>' +
        '<dataRef ref="A1:B2" sheet="average" r:id="rId1"/></dataRefs>',
    );
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

    it("round-trips pageSetup paperWidth/paperHeight with UniversalMeasure (mm)", () => {
      const xml = buildWorksheetXml(
        { pageSetup: { paperWidth: "210mm", paperHeight: "297mm" }, rows: [] },
        {},
      );
      expect(xml).toContain('paperWidth="210mm"');
      expect(xml).toContain('paperHeight="297mm"');
      const el = parseXml(xml).elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      const result = worksheetDesc.parse(el, {} as unknown as ReadContext);
      expect(result.pageSetup?.paperWidth).toBe("210mm");
      expect(result.pageSetup?.paperHeight).toBe("297mm");
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
    it("omits tabSelected unless explicitly set", () => {
      // Only the active sheet carries tabSelected (Excel uses workbookView
      // activeTab); omit it by default so multi-sheet output doesn't mark
      // every sheet active.
      const xml = buildWorksheetXml({ rows: [{ cells: [{ value: "A" }] }] }, {});
      expect(xml).not.toContain("tabSelected");
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
          autoFilter: { ref: "A1:D10", columns: [{ colId: 2, top10: { val: 5 } }] },
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
            columns: [
              { colId: 1, customFilters: { entries: [{ operator: "greaterThan", val: "100" }] } },
            ],
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
          autoFilter: {
            ref: "A1:D10",
            sortState: { ref: "A1:D10", conditions: [{ ref: "B1", descending: true }] },
          },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<sortState ref="A1:D10">');
      expect(xml).toContain('ref="B1"');
      expect(xml).toContain('descending="1"');
    });

    it("generates sheet-level sortState without autoFilter", () => {
      const xml = buildWorksheetXml(
        {
          sortState: { ref: "A1:D10", conditions: [{ ref: "B1", descending: true }] },
          rows: [{ cells: [{ value: "A" }] }],
        },
        {},
      );
      expect(xml).toContain('<sortState ref="A1:D10">');
    });
  });

  describe("autoFilter round-trip", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
      sharedStrings: [],
    } as unknown as ReadContext;

    function roundTrip(opts: WorksheetOptions) {
      const xml = buildWorksheetXml(opts, {});
      // nativeTypeAttributes mirrors the real xlsx parse path (ParsedArchive.get
      // coerces "1"/"0" to numbers), so boolean reads are exercised under coercion.
      const doc = parseXml(xml, { nativeTypeAttributes: true });
      const el = doc.elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("round-trips ref-only autoFilter as shorthand string", () => {
      const result = roundTrip({
        autoFilter: "A1:D10",
        rows: [{ cells: [{ value: "A" }] }],
      });
      expect(result.autoFilter).toBe("A1:D10");
    });

    it("round-trips autoFilter and dataValidation revision UIDs", () => {
      const result = roundTrip({
        autoFilter: { ref: "A1:D10", uid: "{11111111-1111-1111-1111-111111111111}" },
        dataValidations: [
          {
            sqref: "B1:B10",
            type: "whole",
            operator: "between",
            formula1: "1",
            formula2: "10",
            uid: "{22222222-2222-2222-2222-222222222222}",
          },
        ],
        rows: [{ cells: [{ value: "A" }] }],
      });
      expect(result.autoFilter).toMatchObject({
        uid: "{11111111-1111-1111-1111-111111111111}",
      });
      expect(result.dataValidations?.[0]?.uid).toBe("{22222222-2222-2222-2222-222222222222}");
    });

    it("round-trips top10 filter", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [{ colId: 2, top10: { val: 5, percent: true } }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([{ colId: 2, top10: { val: 5, percent: true } }]);
    });

    it("round-trips customFilters with two operands joined by AND", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [
            {
              colId: 1,
              customFilters: {
                and: true,
                entries: [
                  { operator: "greaterThan", val: "100" },
                  { operator: "lessThan", val: "200" },
                ],
              },
            },
          ],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([
        {
          colId: 1,
          customFilters: {
            and: true,
            entries: [
              { operator: "greaterThan", val: "100" },
              { operator: "lessThan", val: "200" },
            ],
          },
        },
      ]);
    });

    it("round-trips filters with values", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [{ colId: 0, filters: { values: ["a", "b"] } }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([{ colId: 0, filters: { values: ["a", "b"] } }]);
    });

    it("round-trips dateGroupItems wrapped in filters", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [
            {
              colId: 2,
              filters: { dateGroupItems: [{ dateTimeGrouping: "year", year: 2020, month: 3 }] },
            },
          ],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([
        {
          colId: 2,
          filters: { dateGroupItems: [{ dateTimeGrouping: "year", year: 2020, month: 3 }] },
        },
      ]);
    });

    it("round-trips colorFilter", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [{ colId: 1, colorFilter: { dxfId: 2, cellColor: false } }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([{ colId: 1, colorFilter: { dxfId: 2, cellColor: false } }]);
    });

    it("round-trips iconFilter", () => {
      // @iconSet is ST_IconSetType — a string enum, never a number.
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [{ colId: 3, iconFilter: { iconSet: "4Arrows", iconId: 2 } }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([{ colId: 3, iconFilter: { iconSet: "4Arrows", iconId: 2 } }]);
    });

    it("round-trips dynamicFilter", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          columns: [{ colId: 0, dynamicFilter: { type: "aboveAverage", val: 42 } }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.columns).toEqual([{ colId: 0, dynamicFilter: { type: "aboveAverage", val: 42 } }]);
    });

    it("round-trips sortState with conditions", () => {
      const result = roundTrip({
        autoFilter: {
          ref: "A1:D10",
          sortState: {
            ref: "A1:D10",
            caseSensitive: true,
            sortMethod: "pinYin",
            conditions: [{ ref: "B1", descending: true }],
          },
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      const af = result.autoFilter as AutoFilterOptions;
      expect(af.sortState).toEqual({
        ref: "A1:D10",
        caseSensitive: true,
        sortMethod: "pinYin",
        conditions: [{ ref: "B1", descending: true }],
      });
    });

    it("round-trips sheet-level sortState separately from autoFilter", () => {
      const result = roundTrip({
        autoFilter: "A1:D10",
        sortState: {
          ref: "A1:D10",
          conditions: [{ ref: "C1", sortBy: "cellColor", dxfId: 2 }],
        },
        rows: [{ cells: [{ value: "A" }] }],
      });
      expect(result.autoFilter).toBe("A1:D10");
      expect(result.sortState).toEqual({
        ref: "A1:D10",
        conditions: [{ ref: "C1", sortBy: "cellColor", dxfId: 2 }],
      });
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
      const el = doc.elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
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
      const el = doc.elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("round-trips outlinePr summaryBelow/summaryRight", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        columns: [{ min: 1, max: 1, outlineLevel: 1 }],
        properties: { outlineSummaryBelow: false, outlineSummaryRight: false },
      };
      const result = roundTrip(opts);
      const sp = result.properties!;

      expect(sp.outlineSummaryBelow).toBe(false);
      expect(sp.outlineSummaryRight).toBe(false);
    });

    it("round-trips explicit outlinePr defaults", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        properties: { outlineSummaryBelow: true, outlineSummaryRight: true },
      });

      expect(result.properties?.outlineSummaryBelow).toBe(true);
      expect(result.properties?.outlineSummaryRight).toBe(true);
    });

    it("round-trips explicit printOptions defaults", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        printOptions: {
          gridLines: true,
          headings: false,
          horizontalCentered: true,
          verticalCentered: false,
        },
      });

      expect(result.printOptions).toEqual({
        gridLines: true,
        headings: false,
        horizontalCentered: true,
        verticalCentered: false,
      });
    });

    it("round-trips explicit pageSetup defaults", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        pageSetup: {
          orientation: "default",
          pageOrder: "downThenOver",
          useFirstPageNumber: false,
          usePrinterDefaults: true,
          blackAndWhite: false,
          draft: false,
          cellComments: "none",
          errors: "displayed",
        },
      });

      expect(result.pageSetup).toEqual({
        orientation: "default",
        pageOrder: "downThenOver",
        useFirstPageNumber: false,
        usePrinterDefaults: true,
        blackAndWhite: false,
        draft: false,
        cellComments: "none",
        errors: "displayed",
      });
    });

    it("round-trips explicit sheetFormat defaults and sparse pageMargins", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        sheetFormat: {
          defaultRowHeight: 15,
          customHeight: false,
          zeroHeight: false,
          dyDescent: 0.25,
        },
        pageMargins: { top: 1, bottom: 1 },
      });

      expect(result.sheetFormat).toEqual({
        defaultRowHeight: 15,
        customHeight: false,
        zeroHeight: false,
        dyDescent: 0.25,
      });
      expect(result.pageMargins).toEqual({
        top: 1,
        bottom: 1,
        topRaw: "1",
        bottomRaw: "1",
      });
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

  describe("rowBreaks/colBreaks round-trip", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
      sharedStrings: [],
    } as unknown as ReadContext;

    function roundTrip(opts: WorksheetOptions) {
      const xml = buildWorksheetXml(opts, {});
      const doc = parseXml(xml);
      const el = doc.elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("round-trips row and column page breaks", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        rowBreaks: [
          { id: 5, manual: true },
          { id: 10, min: 1, max: 100 },
        ],
        colBreaks: [{ id: 3, manual: true }],
      };
      const result = roundTrip(opts);
      expect(result.rowBreaks).toEqual([
        { id: 5, manual: true },
        { id: 10, min: 1, max: 100 },
      ]);
      expect(result.colBreaks).toEqual([{ id: 3, manual: true }]);
    });

    it("round-trips customProperties, cellWatches, legacyDrawingHF", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        customProperties: [
          { name: "MyProp", rId: "rId1" },
          { name: "Other", rId: "rId2" },
        ],
        cellWatches: [{ reference: "A1" }, { reference: "B2" }],
        legacyDrawingHF: "rId3",
      };
      const result = roundTrip(opts);
      expect(result.customProperties).toEqual([
        { name: "MyProp", rId: "rId1" },
        { name: "Other", rId: "rId2" },
      ]);
      expect(result.cellWatches).toEqual([{ reference: "A1" }, { reference: "B2" }]);
      expect(result.legacyDrawingHF).toBe("rId3");
    });

    it("round-trips dataConsolidate", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        dataConsolidate: {
          function: "average",
          topLabels: true,
          leftLabels: true,
          link: true,
          refs: ["Sheet1!A1:B2", "Sheet2!A1:B2"],
        },
      };
      const result = roundTrip(opts);
      expect(result.dataConsolidate).toEqual({
        function: "average",
        topLabels: true,
        leftLabels: true,
        link: true,
        refs: [{ ref: "Sheet1!A1:B2" }, { ref: "Sheet2!A1:B2" }],
      });
    });

    it("round-trips scenarios (what-if analysis)", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        scenarios: {
          current: 0,
          scenarios: [
            {
              name: "Best Case",
              count: 2,
              user: "analyst",
              comment: "Optimistic forecast",
              inputCells: [
                { reference: "B2", val: 100 },
                { reference: "B3", val: "q1" },
              ],
            },
          ],
        },
      };
      const result = roundTrip(opts);
      expect(result.scenarios).toEqual({
        current: 0,
        scenarios: [
          {
            name: "Best Case",
            count: 2,
            user: "analyst",
            comment: "Optimistic forecast",
            inputCells: [
              { reference: "B2", val: 100 },
              { reference: "B3", val: "q1" },
            ],
          },
        ],
      });
    });

    it("round-trips customSheetViews", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        customSheetViews: [
          {
            guid: "A1B2C3",
            scale: 90,
            showPageBreaks: true,
            showGridLines: false,
            showRowColHeaders: false,
            fitToPage: true,
            filter: true,
            state: "hidden",
            view: "pageBreakPreview",
            pageMargins: {
              left: 0.7,
              right: 0.7,
              top: 0.75,
              bottom: 0.75,
              header: 0.3,
              footer: 0.3,
            },
            printOptions: { horizontalCentered: true },
            pageSetup: { orientation: "landscape", scale: 80 },
            headerFooter: { oddHeader: "H", oddFooter: "F" },
          },
        ],
      };
      const result = roundTrip(opts);
      expect(result.customSheetViews).toEqual([
        {
          guid: "A1B2C3",
          scale: 90,
          showPageBreaks: true,
          showGridLines: false,
          showRowColHeaders: false,
          fitToPage: true,
          filter: true,
          state: "hidden",
          view: "pageBreakPreview",
          pageMargins: {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
          },
          printOptions: { horizontalCentered: true },
          pageSetup: { orientation: "landscape", scale: 80 },
          headerFooter: { oddHeader: "H", oddFooter: "F" },
        },
      ]);
    });

    it("round-trips oleObjects and controls", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        oleObjects: [
          {
            shapeId: 1,
            progId: "Excel.Sheet.12",
            rId: "rId1",
            properties: { locked: false, print: false, disabled: true, macro: "Module1.Run" },
          },
        ],
        controls: [{ shapeId: 2, rId: "rId2", name: "Button1", locked: false, linkedCell: "A1" }],
      };
      const result = roundTrip(opts);
      expect(result.oleObjects).toEqual([
        {
          shapeId: 1,
          progId: "Excel.Sheet.12",
          rId: "rId1",
          properties: { locked: false, print: false, disabled: true, macro: "Module1.Run" },
        },
      ]);
      expect(result.controls).toEqual([
        { shapeId: 2, rId: "rId2", name: "Button1", locked: false, linkedCell: "A1" },
      ]);
    });

    it("round-trips mc:AlternateContent wrappers with objectPr anchors", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        oleObjects: [
          {
            shapeId: 1,
            progId: "Paint.Picture",
            rId: "rId4",
            alternateContent: true,
            properties: {
              defaultSize: false,
              iconRid: "rId5",
              anchor: {
                moveWithCells: true,
                from: { col: 0, row: 0 },
                to: { col: 3, colOff: 552450, row: 12, rowOff: 95250 },
              },
            },
          },
        ],
        controls: [
          {
            shapeId: 2,
            rId: "rId6",
            name: "Button1",
            alternateContent: true,
            defaultSize: false,
            autoLine: false,
            autoPict: false,
            iconRid: "rId7",
            anchor: {
              from: { col: 0, row: 0 },
              to: { col: 1, row: 1 },
            },
          },
        ],
      };
      const xml = buildWorksheetXml(opts, {});
      expect(xml).toContain(
        '<mc:Choice xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" Requires="x14">',
      );
      expect(xml).toContain(
        '<mc:Fallback><oleObject shapeId="1" progId="Paint.Picture" r:id="rId4"/></mc:Fallback>',
      );
      expect(xml).toContain(
        '<anchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" moveWithCells="1">',
      );
      expect(xml).toContain("<xdr:col>3</xdr:col><xdr:colOff>552450</xdr:colOff>");
      expect(xml).toContain('autoLine="0" autoPict="0" r:id="rId7"');
      const doc = parseXml(xml, { nativeTypeAttributes: true });
      const el = doc.elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      const result = worksheetDesc.parse(
        el,
        {} as unknown as ReadContext,
      ) as unknown as WorksheetOptions;
      // CT_Marker's four elements are all required in the XSD, so stringify
      // materializes omitted offsets as explicit zeros and parse reads them
      // back — compare against the materialized form.
      expect(result.oleObjects).toEqual([
        {
          ...opts.oleObjects![0]!,
          properties: {
            ...opts.oleObjects![0]!.properties!,
            anchor: {
              ...opts.oleObjects![0]!.properties!.anchor!,
              from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
            },
          },
        },
      ]);
      expect(result.controls).toEqual([
        {
          ...opts.controls![0]!,
          anchor: {
            from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
            to: { col: 1, colOff: 0, row: 1, rowOff: 0 },
          },
        },
      ]);
    });

    it("round-trips webPublishItems and ext", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        webPublishItems: [
          {
            id: 1,
            divId: "div1",
            sourceType: "range",
            destinationFile: "out.htm",
            title: "Report",
            autoRepublish: true,
          },
        ],
        ext: '<ext uri="{xxx}"><x:foo xmlns:x="x"/></ext>',
      };
      const result = roundTrip(opts);
      expect(result.webPublishItems).toEqual([
        {
          id: 1,
          divId: "div1",
          sourceType: "range",
          destinationFile: "out.htm",
          title: "Report",
          autoRepublish: true,
        },
      ]);
      expect(result.ext).toContain('<ext uri="{xxx}">');
      expect(result.ext).toContain("x:foo");
    });

    it("round-trips drawingHF (header/footer drawing offsets)", () => {
      const opts: WorksheetOptions = {
        rows: [{ cells: [{ value: "A" }] }],
        drawingHF: { rId: "rId5", lho: 3, che: 4, rff: 5 },
      };
      const result = roundTrip(opts);
      expect(result.drawingHF).toEqual({ rId: "rId5", lho: 3, che: 4, rff: 5 });
    });
  });

  describe("default-true attributes and legacy password round-trip", () => {
    const readCtx = {
      resolveRelationship: () => undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
      sharedStrings: [],
    } as unknown as ReadContext;

    function roundTrip(opts: WorksheetOptions) {
      const xml = buildWorksheetXml(opts, {});
      const doc = parseXml(xml, { nativeTypeAttributes: true });
      const el = doc.elements?.[0];
      if (!el) throw new Error("parsed document has no root element");
      return worksheetDesc.parse(el, readCtx) as unknown as WorksheetOptions;
    }

    it("round-trips sheetPr published/enableFormatConditionsCalculation = false (XSD default true)", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        properties: { published: false, enableFormatConditionsCalculation: false },
      });
      expect(result.properties).toEqual({
        published: false,
        enableFormatConditionsCalculation: false,
      });
    });

    it("round-trips pageSetup usePrinterDefaults = false (XSD default true)", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        pageSetup: { usePrinterDefaults: false },
      });
      expect(result.pageSetup?.usePrinterDefaults).toBe(false);
    });

    it("keeps a headerFooter element that carries only false-valued flags", () => {
      const xml = buildWorksheetXml(
        { rows: [{ cells: [{ value: "A" }] }], headerFooter: { scaleWithDoc: false } },
        {},
      );
      expect(xml).toContain('<headerFooter scaleWithDoc="0"/>');
    });

    it("does not read back the protectedRange legacy @password hash", () => {
      const result = roundTrip({
        rows: [{ cells: [{ value: "A" }] }],
        protectedRanges: [{ sqref: "A1:C10", name: "Range1", password: "secret" }],
      });
      expect(result.protectedRanges?.[0]?.password).toBeUndefined();
    });
  });
});
