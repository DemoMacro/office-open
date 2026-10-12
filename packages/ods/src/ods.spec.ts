import { generateOcf, ODF_NAMESPACES, readOcf } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook, OdsParseError, parseWorkbook, type OdsDocumentOptions } from "./index";

describe("ODS codec", () => {
  it("round-trips a workbook through canonical WorkbookOptions", () => {
    const parsed = parseWorkbook(
      generateWorkbook({ worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells).toEqual([{ reference: "A1", value: "ODS" }]);
  });

  it("round-trips percentage, currency, and time cell semantics", () => {
    const parsed = parseWorkbook(
      generateWorkbook({
        worksheets: [
          {
            rows: [
              {
                cells: [
                  { value: 0.25, style: { numFmt: "0%" } },
                  { value: 1.25, style: { numFmt: "[$$]#,##0.00" } },
                  { value: 0.5, style: { numFmt: "hh:mm:ss" } },
                ],
              },
            ],
          },
        ],
      }),
    );
    const cells = parsed.worksheets?.[0]?.rows?.[0]?.cells ?? [];
    expect(cells[0]).toEqual({ reference: "A1", value: 0.25, style: { numFmt: "0%" } });
    expect(cells[1]).toEqual({
      reference: "B1",
      value: 1.25,
      style: { numFmt: "[$$]#,##0.00" },
    });
    expect(cells[2]).toEqual({
      reference: "C1",
      value: 0.5,
      style: { numFmt: "hh:mm:ss" },
    });
  });

  it("wraps invalid packages in OdsParseError", () => {
    expect(() => parseWorkbook(new Uint8Array([1, 2, 3]))).toThrow(OdsParseError);
  });

  it("preserves unstyled percentage and currency values", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="percentage" office:value="0.25"/><table:table-cell office:value-type="currency" office:value="12.5" office:currency="USD"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    ).worksheets?.[0];
    if (!worksheet) throw new Error("Missing worksheet");
    expect(worksheet.rows?.[0]?.cells?.[0]).toMatchObject({
      value: 0.25,
      style: { numFmt: "0%" },
    });
    expect(worksheet.rows?.[0]?.cells?.[1]).toMatchObject({
      value: 12.5,
      style: { numFmt: "[$USD] #,##0.00" },
    });
  });

  it("preserves negative time durations", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="time" office:time-value="-PT1H"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const parsed = parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    ).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value;
    expect(parsed).toBe(-1 / 24);

    const roundTrip = parseWorkbook(
      generateWorkbook({
        worksheets: [{ rows: [{ cells: [{ value: -1 / 24, style: { numFmt: "[hh]:mm:ss" } }] }] }],
      }),
    );
    expect(
      parseWorkbook(generateWorkbook(roundTrip)).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value,
    ).toBe(-1 / 24);
  });

  it("expands column, row, and cell repeats", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="3"/><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="3"/><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    ).worksheets?.[0];
    if (!worksheet) throw new Error("Missing worksheet");
    expect(worksheet.columns).toEqual([{ min: 1, max: 3, width: undefined, hidden: undefined }]);
    expect(worksheet.rows).toHaveLength(2);
    expect(worksheet.rows?.[0]?.rowNumber).toBe(1);
    expect(worksheet.rows?.[1]?.rowNumber).toBe(2);
    expect(worksheet.rows?.[0]?.cells?.map((cell) => cell.reference)).toEqual([
      "A1",
      "B1",
      "C1",
      "D1",
    ]);
    expect(worksheet.rows?.[0]?.cells?.[3]?.value).toBe(1);
  });

  it("does not materialize very large empty repeats", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="999999999"><table:table-cell table:number-columns-repeated="16384"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    ).worksheets?.[0];
    expect(worksheet?.rows).toEqual([]);
  });

  it("compresses very large empty row repeats into the default row height", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="15pt"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:style-name="ro1" table:number-rows-repeated="1048576"><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    ).worksheets?.[0];
    expect(worksheet?.rows).toEqual([]);
    expect(worksheet?.sheetFormat).toEqual({ defaultRowHeight: 15 });
  });

  it("round-trips merged cells through spans and covered cells", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2" office:value-type="string"><text:p>Merged</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const source = parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    );
    expect(source.worksheets?.[0]?.mergeCells).toEqual([{ ref: "A1:B2" }]);

    const roundTrip = parseWorkbook(generateWorkbook(source));
    expect(roundTrip.worksheets?.[0]?.mergeCells).toEqual([{ ref: "A1:B2" }]);
    expect(roundTrip.worksheets?.[0]?.rows?.[0]?.cells?.map((cell) => cell.reference)).toEqual([
      "A1",
      "B1",
    ]);
    expect(roundTrip.worksheets?.[0]?.rows?.[1]?.cells?.map((cell) => cell.reference)).toEqual([
      "A2",
      "B2",
    ]);
  });

  it("requires an embedded object body but tolerates a missing replacement", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell><draw:frame svg:width="1cm" svg:height="1cm"><draw:object-ole xlink:href="./Object 1" xlink:type="simple"/></draw:frame></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": content,
      "Object 1": new Uint8Array([1]),
    });
    const parsed = parseWorkbook(source) as OdsDocumentOptions;
    expect(parsed.packageMembers).toEqual([
      { path: "Object 1", mediaType: "application/binary", data: new Uint8Array([1]) },
    ]);
    expect(parseWorkbook(generateWorkbook(parsed)).packageMembers).toEqual(parsed.packageMembers);

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
    expect(error).toMatchObject({ reason: "embedded-object-subdocument-missing" });
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

  it("emits linked-only cell images without a package media part", () => {
    const image = {
      type: "image" as const,
      href: "https://example.com/logo.png",
      sourceUrl: "https://example.com/logo.png",
      width: 914400,
      height: 914400,
    };
    const generated = generateWorkbook({
      worksheets: [{ rows: [{ cells: [{ value: "Logo", graphics: [image] }] }] }],
    });
    const { binaries, manifest } = readOcf(
      generated,
      "application/vnd.oasis.opendocument.spreadsheet",
    );

    expect(Object.keys(binaries)).toEqual([]);
    expect(manifest.entries.map((entry) => entry.fullPath)).not.toContain(image.href);
    expect(
      parseWorkbook(generated).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics?.[0],
    ).toMatchObject({ sourceUrl: image.sourceUrl, data: undefined });
  });

  it("restores percent and length semantics in style overlays", () => {
    const options: OdsDocumentOptions = {
      worksheets: [{ rows: [{ cells: [{ value: "Overlay" }] }] }],
      styleOverlays: [
        {
          name: "ro1",
          family: "table-row",
          properties: [
            {
              name: "style:table-row-properties",
              attributes: { "style:row-height": 20, "style:use-optimal-row-height": false },
            },
          ],
        },
        {
          name: "P1",
          family: "paragraph",
          properties: [
            {
              name: "style:paragraph-properties",
              attributes: { "fo:line-height": 150 },
            },
          ],
        },
      ],
    };
    const generated = generateWorkbook(options);
    const content = readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet").files[
      "content.xml"
    ]!;
    expect(content).toContain('style:row-height="20pt"');
    expect(content).toContain('fo:line-height="150%"');
    expect(content).not.toMatch(/(?:row-height|line-height)="(?:150|20)"(?:\s|\/>)/);
  });

  it("rejects a cell image with neither embedded data nor a source URL", () => {
    let error: unknown;
    try {
      generateWorkbook({
        worksheets: [
          {
            rows: [
              {
                cells: [
                  { value: "Logo", graphics: [{ type: "image", href: "Pictures/logo.png" }] },
                ],
              },
            ],
          },
        ],
      });
    } catch (cause) {
      error = cause;
    }

    expect(error).toBeInstanceOf(OdsParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      name: "draw:image",
      reason: "missing-image-source",
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
