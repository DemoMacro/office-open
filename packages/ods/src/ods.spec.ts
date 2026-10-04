import { generateOcf, ODF_NAMESPACES } from "@office-open/ocf";
import { describe, expect, it } from "vite-plus/test";

import { generateOds, OdsParseError, parseOds, type OdsWorkbookOptions } from "./index";
import type { OdsCellOptions } from "./semantics";

describe("ODS codec", () => {
  it("round-trips a workbook through canonical WorkbookOptions", () => {
    const parsed = parseOds(
      generateOds({ worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells).toEqual([{ reference: "A1", value: "ODS" }]);
  });

  it("wraps invalid packages in OdsParseError", () => {
    expect(() => parseOds(new Uint8Array([1, 2, 3]))).toThrow(OdsParseError);
  });

  it("round-trips calculation settings, covered cells, forms, annotations, and cell graphics", () => {
    const options: OdsWorkbookOptions = {
      odfSemantics: {
        calculationSettings: {
          caseSensitive: true,
          nullDate: "1899-12-30",
          iteration: { enabled: true, steps: 20, maximumDifference: 0.001 },
        },
        forms: {
          forms: [{ name: "Sheet form", controls: [{ kind: "form:checkbox", id: "active" }] }],
        },
      },
      worksheets: [
        {
          annotations: [
            { reference: "B2", paragraphs: ["Reviewed"], author: "Author", date: "2026-01-01" },
          ],
          cellGraphics: [
            {
              reference: "B2",
              href: "Pictures/graphic.png",
              name: "Graphic",
              x: 0,
              y: 0,
              width: 360000,
              height: 360000,
            },
          ],
          rows: [
            {
              cells: [
                { reference: "A1", covered: true } as OdsCellOptions,
                { reference: "B2", value: "ODS" },
              ],
            },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(options));
    expect(parsed.odfSemantics).toMatchObject(options.odfSemantics ?? {});
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]).toMatchObject({
      reference: "A1",
      covered: true,
    });
    expect(parsed.worksheets?.[0]?.annotations?.[0]).toMatchObject({
      ...options.worksheets![0]!.annotations![0],
      reference: "B1",
    });
    expect(parsed.worksheets?.[0]?.cellGraphics?.[0]).toMatchObject({
      ...options.worksheets![0]!.cellGraphics![0],
      reference: "B1",
    });
  });

  it("round-trips real chart anchors, semantics, worksheet links, and order", () => {
    const parsed = parseOds(
      generateOds({
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
    expect(charts[0]).not.toHaveProperty("col");
    expect(charts[0]).not.toHaveProperty("row");
  });

  it("round-trips worksheet forms and unreferenced object frames", () => {
    const options: OdsWorkbookOptions = {
      worksheets: [
        {
          forms: {
            forms: [{ name: "Controls", controls: [{ kind: "form:checkbox", id: "active" }] }],
          },
          objectGraphics: [{ reference: "", name: "Embedded", width: 720000, height: 720000 }],
          rows: [{ cells: [{ reference: "A1", value: "ODS" }] }],
        },
      ],
    };
    const parsed = parseOds(generateOds(options));
    expect(parsed.worksheets?.[0]?.forms).toMatchObject(options.worksheets?.[0]?.forms ?? {});
    expect(parsed.worksheets?.[0]?.objectGraphics?.[0]).toMatchObject({
      name: "Embedded",
      width: 720000,
      height: 720000,
    });
  });

  it("round-trips defined names and rejects unknown worksheet children", () => {
    const parsed = parseOds(
      generateOds({
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
      parseOds(
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
});
