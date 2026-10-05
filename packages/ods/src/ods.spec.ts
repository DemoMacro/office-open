import { generateOcf, ODF_NAMESPACES } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook, OdsParseError, parseWorkbook } from "./index";

describe("ODS codec", () => {
  it("round-trips a workbook through canonical WorkbookOptions", () => {
    const parsed = parseWorkbook(
      generateWorkbook({ worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells).toEqual([{ reference: "A1", value: "ODS" }]);
  });

  it("wraps invalid packages in OdsParseError", () => {
    expect(() => parseWorkbook(new Uint8Array([1, 2, 3]))).toThrow(OdsParseError);
  });

  it("round-trips real chart anchors, semantics, worksheet links, and order", () => {
    const parsed = parseWorkbook(
      generateWorkbook({
        worksheets: [
          {
            name: "Revenue",
            rows: [{ cells: [{ value: "ODS" }] }],
            charts: [
              {
                name: "First",
                type: "bar",
                series: [{ name: "Sales", values: [1, 2] }],
                col: 1,
                row: 1,
                anchorType: "absolute",
                absoluteX: 360000,
                absoluteY: 720000,
                extentCx: 2160000,
                extentCy: 1440000,
                showLegend: true,
                legendPosition: "top",
              },
              {
                name: "Second",
                type: "scatter",
                series: [{ name: "Points", xValues: [1, 2], yValues: [2, 4] }],
                col: 1,
                row: 1,
                anchorType: "absolute",
                absoluteX: 1080000,
                absoluteY: 1440000,
                extentCx: 1800000,
                extentCy: 1080000,
              },
            ],
          },
        ],
      }),
    );
    const charts = parsed.worksheets?.[0]?.charts ?? [];
    expect(charts.map((chart) => chart.name)).toEqual(["First", "Second"]);
    expect(charts[0]).toMatchObject({
      type: "bar",
      anchorType: "absolute",
      absoluteX: 360000,
      absoluteY: 720000,
      extentCx: 2160000,
      extentCy: 1440000,
      showLegend: true,
      legendPosition: "top",
      series: [{ name: "Sales", values: [1, 2] }],
    });
    expect(charts[1]).toMatchObject({
      type: "scatter",
      series: [{ name: "Points", xValues: [1, 2], yValues: [2, 4] }],
    });
  });

  it("round-trips defined names and rejects unknown worksheet children", () => {
    const parsed = parseWorkbook(
      generateWorkbook({
        definedNames: [
          { name: "Total", value: "Revenue!$A$1:$A$2" },
          { name: "Double", value: "SUM(Revenue!$A$1:$A$2)*2" },
        ],
        worksheets: [{ name: "Revenue", rows: [] }],
      }),
    );
    expect(parsed.definedNames).toEqual([
      { name: "Total", value: "Revenue!$A$1:$A$2" },
      { name: "Double", value: "SUM(Revenue!$A$1:$A$2)*2" },
    ]);

    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:unknown/></table:table></office:spreadsheet></office:body></office:document-content>`;
    let error: unknown;
    try {
      parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdsParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      path: '/office:document-content/office:body/office:spreadsheet/table:table[@table:name="Sheet1"]/table:unknown',
      name: "table:unknown",
      reason: "element has no canonical WorksheetOptions mapping",
    });
  });

  it("round-trips cell object frames and rejects unknown frame children", () => {
    const object = {
      type: "object" as const,
      href: "embedded-object",
      name: "Legacy object",
      description: "Embedded legacy object",
      x: 360000,
      y: 720000,
      width: 2160000,
      height: 1440000,
    };
    const parsed = parseWorkbook(
      generateWorkbook({ worksheets: [{ rows: [{ cells: [{ value: 1, graphics: [object] }] }] }] }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics).toEqual([object]);

    const oleContent = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value="1"><draw:frame draw:name="Legacy object" svg:x="1cm" svg:y="2cm" svg:width="6cm" svg:height="4cm"><draw:object-ole xlink:href="./embedded-object" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><svg:desc>Embedded legacy object</svg:desc><draw:unknown/></draw:frame></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    let error: unknown;
    try {
      parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
          "content.xml": oleContent,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdsParseError);
    expect(error).toMatchObject({ part: "content.xml", name: "draw:unknown" });
  });

  it("round-trips calculation settings, forms, and covered cells", () => {
    const calculation = {
      iterate: true,
      iterateCount: 5,
      iterateDelta: 0.01,
      caseSensitive: false,
      precisionAsShown: true,
      automaticLabelSearch: true,
      regularExpressionSearch: true,
      wildcardSearch: true,
      nullDate: { year: 1900, month: 1, day: 1 },
    };
    const forms = [
      {
        name: "Inputs",
        automaticFocus: true,
        designMode: true,
        controls: [
          { control: "text" as const, id: "text", name: "Name", value: "Value", maxLength: 24 },
          { control: "checkBox" as const, id: "check", checked: true },
          {
            control: "dropDownList" as const,
            id: "list",
            entries: ["One", "Two"],
            selectedIndex: 1,
          },
        ],
      },
    ];
    const parsed = parseWorkbook(
      generateWorkbook({
        calculation,
        forms,
        worksheets: [
          {
            rows: [{ cells: [{ value: 1 }, { covered: { reference: "B1", text: "covered" } }] }],
          },
        ],
      }),
    );
    expect(parsed.calculation).toEqual(calculation);
    expect(parsed.forms).toEqual(forms);
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[1]?.covered).toEqual({
      reference: "B1",
      text: "covered",
    });
  });

  it("rejects unknown calculation-setting children", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:calculation-settings><table:unknown/></table:calculation-settings></office:spreadsheet></office:body></office:document-content>`;
    let error: unknown;
    try {
      parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdsParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      name: "table:unknown",
      reason: "element has no canonical calculation mapping",
    });
  });
});
