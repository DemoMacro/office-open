import type { FormControlOptions, OfficeFormsOptions } from "@office-open/core";
import { generateOcf, ODF_NAMESPACES, readOcf } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import {
  generateWorkbook,
  OdsParseError,
  parseWorkbook,
  type OdsWorkbookOptions,
  type WorksheetOptions,
} from "./index";

describe("ODS codec", () => {
  it("projects and regenerates graphic gradient fills", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="Gradient" style:family="graphic"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Blue"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:rect draw:style-name="Gradient" svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="2cm"/></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
      "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles><draw:gradient draw:name="Blue" draw:style="linear" draw:start-color="#112233" draw:end-color="#AABBCC"/></office:styles></office:document-styles>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.shapes?.[0]?.properties?.fill).toMatchObject({
      type: "gradient",
      stops: [
        { position: 0, color: "112233" },
        { position: 100, color: "AABBCC" },
      ],
    });
    const generated = await generateWorkbook(parsed);
    const files = (await readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet"))
      .files;
    expect(files["styles.xml"]).toContain('draw:gradient draw:name="gradient-1"');
    const reread = await parseWorkbook(generated);
    expect(reread.worksheets?.[0]?.shapes?.[0]?.properties?.fill).toMatchObject({
      type: "gradient",
    });
  });

  it("rejects a content document without its required spreadsheet body", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body/></office:document-content>`,
    });

    await expect(parseWorkbook(source)).rejects.toMatchObject({
      name: "OdsParseError",
      part: "content.xml",
      elementName: "office:spreadsheet",
      reason: "missing-required-element",
    });
  });

  it("round-trips a workbook through canonical WorkbookOptions", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({ worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells).toEqual([{ reference: "A1", value: "ODS" }]);
  });

  it("round-trips grouped worksheet shapes through canonical groups", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:rect draw:name="Standalone" svg:x="5cm" svg:y="6cm" svg:width="2cm" svg:height="1cm"/><draw:g draw:name="Group" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><svg:title>Group title</svg:title><draw:rect draw:name="Child" svg:x="0.5cm" svg:y="0.25cm" svg:width="2cm" svg:height="1cm"><text:p>Grouped text</text:p></draw:rect><draw:g draw:name="Nested" svg:x="0.25cm" svg:y="0.5cm" svg:width="1cm" svg:height="1cm"><draw:ellipse draw:name="Nested child" svg:x="0.1cm" svg:y="0.1cm" svg:width="0.5cm" svg:height="0.5cm"/></draw:g></draw:g></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.shapes).toMatchObject([
      { name: "Standalone", anchorType: "absolute", absoluteX: 1800000, absoluteY: 2160000 },
    ]);
    expect(parsed.worksheets?.[0]?.groups).toMatchObject([
      {
        col: 1,
        row: 1,
        anchorType: "absolute",
        absoluteX: 360000,
        absoluteY: 720000,
        extentCx: 1440000,
        extentCy: 1080000,
        name: "Group",
        title: "Group title",
        properties: { x: 360000, y: 720000, width: 1440000, height: 1080000 },
        shapes: [
          {
            name: "Child",
            properties: { x: 180000, y: 90000, width: 720000, height: 360000 },
            textBody: { text: "Grouped text" },
          },
        ],
        groups: [
          {
            name: "Nested",
            properties: { x: 90000, y: 180000, width: 360000, height: 360000 },
            shapes: [
              {
                name: "Nested child",
                properties: { x: 36000, y: 36000, width: 180000, height: 180000 },
              },
            ],
          },
        ],
      },
    ]);

    const generated = await generateWorkbook(parsed);
    const roundTripped = await parseWorkbook(generated);
    expect(roundTripped.worksheets?.[0]?.groups).toEqual(parsed.worksheets?.[0]?.groups);
  });

  it("projects worksheet group rotation into canonical group properties", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:g draw:name="Rotated" draw:transform="rotate(15deg)" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><draw:rect svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/></draw:g></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.groups?.[0]?.properties.rotation).toBe(15);

    const roundTripped = await parseWorkbook(await generateWorkbook(parsed));
    expect(roundTripped.worksheets?.[0]?.groups?.[0]?.properties.rotation).toBe(15);
  });

  it("round-trips worksheet shape anchors, rotation, and end-cell offsets", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:rect draw:name="Box" draw:transform="rotate(15deg)" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm" table:end-cell-address="C4" table:end-x="5mm" table:end-y="6mm"/></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.shapes?.[0]).toMatchObject({
      toCol: 3,
      toRow: 4,
      toColOffset: 180000,
      toRowOffset: 216000,
      properties: { rotation: 15 },
    });

    const reread = await parseWorkbook(await generateWorkbook(parsed));
    expect(reread.worksheets?.[0]?.shapes?.[0]).toMatchObject({
      toCol: 3,
      toRow: 4,
      toColOffset: 180000,
      toRowOffset: 216000,
      properties: { rotation: 15 },
    });
  });

  it("projects worksheet-level image frames into WorkbookOptions", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:frame draw:name="Logo" draw:transform="rotate(15deg)" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm" table:end-cell-address="C4" table:end-x="5mm" table:end-y="6mm"><draw:image xlink:href="Pictures/logo.png"/><svg:title>Logo title</svg:title><svg:desc>Logo description</svg:desc></draw:frame></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
      "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
      "Pictures/logo.png": new Uint8Array([1, 2, 3]),
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.images?.[0]).toMatchObject({
      type: "png",
      sourcePath: "Pictures/logo.png",
      anchorType: "twoCell",
      col: 1,
      row: 1,
      colOffset: 360000,
      rowOffset: 720000,
      toCol: 3,
      toRow: 4,
      toColOffset: 180000,
      toRowOffset: 216000,
      properties: {
        width: 1440000,
        height: 1080000,
        rotation: 15,
      },
    });
    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]?.images).toEqual(parsed.worksheets?.[0]?.images);
  });

  it("projects worksheet group fills into canonical group properties", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="GroupStyle" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#3366ff"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:g draw:name="Filled" draw:style-name="GroupStyle" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><draw:rect svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/></draw:g></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.groups?.[0]?.properties.fill).toEqual({
      type: "solid",
      color: "3366ff",
    });

    const roundTripped = await parseWorkbook(await generateWorkbook(parsed));
    expect(roundTripped.worksheets?.[0]?.groups?.[0]?.properties.fill).toEqual(
      parsed.worksheets?.[0]?.groups?.[0]?.properties.fill,
    );
  });

  it("rejects worksheet chart frames whose chart subdocument is missing", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:frame><draw:object xlink:href="Object 1"/></draw:frame></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });

    await expect(parseWorkbook(source)).rejects.toMatchObject({
      name: "OdsParseError",
      part: "content.xml",
      reason: "missing-chart-subdocument",
    });
  });

  it("rejects OOXML pivot definitions that ODS cannot represent", async () => {
    await expect(
      generateWorkbook({
        worksheets: [
          {
            name: "Sheet1",
            rows: [],
            pivotTables: [
              {
                mode: "definition",
                cacheId: 1,
                definition: { name: "Pivot", cacheId: 1 },
              },
            ],
          },
        ],
      }),
    ).rejects.toMatchObject({
      name: "OdsGenerateError",
      reason: "pivot-definition-not-representable",
    });
  });

  it("rejects standalone OOXML pivot caches", async () => {
    await expect(
      generateWorkbook({
        pivotCaches: [
          {
            mode: "definition",
            cacheId: 1,
            definitionPath: "xl/pivotCache/pivotCacheDefinition1.xml",
            definition: {},
          },
        ],
        worksheets: [],
      }),
    ).rejects.toMatchObject({
      name: "OdsGenerateError",
      reason: "pivot-cache-not-representable",
    });
  });

  it("compresses identical adjacent rows without losing reparse positions", async () => {
    const row: WorksheetOptions["rows"] = [{ cells: [{ value: "Shared" }, { value: "Shared" }] }];
    const generated = await generateWorkbook({ worksheets: [{ rows: [...row, ...row, ...row] }] });
    const content = (await readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet"))
      .files["content.xml"]!;
    expect(content).toContain('table:number-columns-repeated="2"');
    expect(content).toContain('table:number-rows-repeated="3"');
    const parsed = await parseWorkbook(generated);
    expect(parsed.worksheets?.[0]?.rows?.map((entry) => entry.rowNumber)).toEqual([1, 2, 3]);
    expect(parsed.worksheets?.[0]?.rows?.[2]?.cells).toEqual([
      { reference: "A3", value: "Shared" },
      { reference: "B3", value: "Shared" },
    ]);
  });

  it("does not bridge nonadjacent identical rows", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [
          {
            rows: [
              { rowNumber: 1, cells: [{ value: "Shared" }] },
              { rowNumber: 3, cells: [{ value: "Shared" }] },
            ],
          },
        ],
      }),
    );
    expect(parsed.worksheets?.[0]?.rows?.map((entry) => entry.rowNumber)).toEqual([1, 2, 3]);
    expect(
      parsed.worksheets?.[0]?.rows?.[1]?.cells?.every((cell) => cell.value === undefined),
    ).toBe(true);
    expect(parsed.worksheets?.[0]?.rows?.[2]?.cells?.[0]?.value).toBe("Shared");
  });

  it("round-trips percentage, currency, and time cell semantics", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
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

  it("round-trips rich cell text through the canonical rich-text model", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [
          {
            rows: [
              {
                cells: [
                  {
                    value: {
                      runs: [
                        { text: "Plain " },
                        {
                          text: "bold",
                          properties: { bold: true, font: "Aptos", size: 12 },
                        },
                        { text: " and " },
                        { text: "red", properties: { color: { rgb: "FFFF0000" } } },
                      ],
                    },
                  },
                ],
              },
            ],
          },
        ],
      }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toEqual({
      runs: [
        { text: "Plain " },
        { text: "bold", properties: { bold: true, font: "Aptos", size: 12 } },
        { text: " and " },
        { text: "red", properties: { color: { rgb: "FFFF0000" } } },
      ],
    });
  });

  it("applies common table-cell styles from styles.xml", async () => {
    const parsed = await parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="CommonCell" office:value-type="string"><text:p>Styled</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
        "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles><style:style style:name="CommonCell" style:family="table-cell"><style:text-properties fo:font-weight="bold" fo:font-size="14pt"/></style:style></office:styles></office:document-styles>`,
      }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.style).toEqual({
      font: { bold: true, size: 14 },
    });
  });

  it("round-trips cell range sources and detective metadata", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [
          {
            rows: [
              {
                cells: [
                  {
                    reference: "A1",
                    cellRangeSource: {
                      name: "Source",
                      lastColumnSpanned: 3,
                      lastRowSpanned: 4,
                      href: "../data.ods#Sheet1.A1:C4",
                      filterName: "Calc MS Excel 2007 XML",
                      filterOptions: "44,34,76,1",
                      refreshDelay: "PT1H",
                    },
                    detective: {
                      highlightedRanges: [
                        {
                          cellRangeAddress: "Sheet1.B2:C3",
                          direction: "fromSameTable",
                          containsError: false,
                        },
                        { markedInvalid: true },
                      ],
                      operations: [{ name: "tracePrecedents", index: 1 }],
                    },
                  },
                ],
              },
            ],
          },
        ],
      }),
    );
    const cell = parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0];
    expect(cell).toMatchObject({
      cellRangeSource: {
        name: "Source",
        lastColumnSpanned: 3,
        lastRowSpanned: 4,
        href: "../data.ods#Sheet1.A1:C4",
        filterName: "Calc MS Excel 2007 XML",
        filterOptions: "44,34,76,1",
        refreshDelay: "PT1H",
      },
      detective: {
        highlightedRanges: [
          {
            cellRangeAddress: "Sheet1.B2:C3",
            direction: "fromSameTable",
            containsError: false,
          },
          { markedInvalid: true },
        ],
        operations: [{ name: "tracePrecedents", index: 1 }],
      },
    });
    const reread = await parseWorkbook(await generateWorkbook(parsed));
    expect(reread.worksheets?.[0]?.rows?.[0]?.cells?.[0]).toEqual(cell);
  });

  it("wraps invalid packages in OdsParseError", async () => {
    await expect(parseWorkbook(new Uint8Array([1, 2, 3]))).rejects.toThrow(OdsParseError);
  });

  it("preserves unstyled percentage and currency values", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="percentage" office:value="0.25"/><table:table-cell office:value-type="currency" office:value="12.5" office:currency="USD"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
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

  it("preserves negative time durations", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="time" office:time-value="-PT1H"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const parsed = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
    ).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value;
    expect(parsed).toBe(-1 / 24);

    const roundTrip = await parseWorkbook(
      await generateWorkbook({
        worksheets: [{ rows: [{ cells: [{ value: -1 / 24, style: { numFmt: "[hh]:mm:ss" } }] }] }],
      }),
    );
    expect(
      (await parseWorkbook(await generateWorkbook(roundTrip))).worksheets?.[0]?.rows?.[0]
        ?.cells?.[0]?.value,
    ).toBe(-1 / 24);
  });

  it("rejects invalid scalar cell values instead of coercing them", async () => {
    const packageFor = (cell: string) =>
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell ${cell}/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
      });
    await expect(
      parseWorkbook(packageFor('office:value-type="time" office:time-value="invalid"')),
    ).rejects.toMatchObject({
      elementName: "office:time-value",
      reason: "invalid time duration",
    });
    await expect(
      parseWorkbook(packageFor('office:value-type="boolean" office:boolean-value="yes"')),
    ).rejects.toMatchObject({
      elementName: "office:boolean-value",
      reason: "invalid boolean value",
    });
    await expect(parseWorkbook(packageFor('office:value-type="future"'))).rejects.toMatchObject({
      elementName: "office:value-type",
      reason: "unsupported value type: future",
    });
  });

  it("expands column, row, and cell repeats", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="3"/><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="3"/><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
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

  it("maps ODF dimension units and boolean lexemes", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="Column" style:family="table-column"><style:table-column-properties style:column-width="2cm"/></style:style><style:style style:name="Row" style:family="table-row"><style:table-row-properties style:row-height="10mm"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-column table:style-name="Column" table:visibility="collapse"/><table:table-row table:style-name="Row" table:visibility="visible"><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.columns?.[0]).toMatchObject({
      width: expect.closeTo((2 * 96) / 2.54, 0.01),
      hidden: true,
    });
    expect(parsed.worksheets?.[0]?.rows?.[0]).toMatchObject({
      height: expect.closeTo((10 * 72) / 25.4, 0.01),
      hidden: undefined,
    });
  });

  it("does not materialize very large empty repeats", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="999999999"><table:table-cell table:number-columns-repeated="16384"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
    ).worksheets?.[0];
    expect(worksheet?.rows).toEqual([]);
  });

  it("clamps virtual default column repeats", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="999999999"/><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
    ).worksheets?.[0];
    expect(worksheet?.columns).toEqual([
      { min: 1, max: 16384, width: undefined, hidden: undefined },
    ]);
    expect(worksheet?.rows?.[0]?.cells?.[0]?.value).toBe(1);
  });

  it("parses cells beyond spreadsheet-specific column limits", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-repeated="20000"/><table:table-cell office:value-type="float" office:value="7"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const parsed = await parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.reference).toBe("ACOG1");
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe(7);
  });

  it("compresses very large empty row repeats into the default row height", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="15pt"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:style-name="ro1" table:number-rows-repeated="1048576"><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const worksheet = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
    ).worksheets?.[0];
    expect(worksheet?.rows).toEqual([]);
    expect(worksheet?.sheetFormat).toEqual({ defaultRowHeight: 15 });
    const regenerated = await parseWorkbook(await generateWorkbook({ worksheets: [worksheet!] }));
    expect(regenerated.worksheets?.[0]?.sheetFormat).toEqual({ defaultRowHeight: 15 });
    expect(regenerated.worksheets?.[0]?.rows).toEqual([]);
  });

  it("preserves default row height alongside explicit rows", async () => {
    const workbook = {
      worksheets: [
        {
          name: "Sheet1",
          sheetFormat: { defaultRowHeight: 18 },
          rows: [{ rowNumber: 1, cells: [{ reference: "A1", value: "value" }] }],
        },
      ],
    } satisfies OdsWorkbookOptions;

    const parsed = await parseWorkbook(await generateWorkbook(workbook));
    expect(parsed.worksheets?.[0]?.sheetFormat).toEqual({ defaultRowHeight: 18 });
    expect(parsed.worksheets?.[0]?.rows?.[0]?.height).toBeUndefined();
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe("value");
  });

  it("round-trips optimal column width", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="Optimal" style:family="table-column"><style:table-column-properties style:use-optimal-column-width="true"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-column table:style-name="Optimal"/><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.columns?.[0]).toMatchObject({ bestFit: true });
    expect(parsed.worksheets?.[0]?.columns?.[0]?.width).toBeUndefined();

    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]?.columns?.[0]).toMatchObject({ bestFit: true });
    expect(regenerated.worksheets?.[0]?.columns?.[0]?.width).toBeUndefined();
  });

  it("round-trips canonical cell font and alignment properties", async () => {
    const workbook = {
      worksheets: [
        {
          name: "Sheet1",
          rows: [
            {
              cells: [
                {
                  reference: "A1",
                  value: "Styled",
                  style: {
                    font: { font: "Liberation Sans" },
                    alignment: { textRotation: 45, shrinkToFit: true },
                  },
                },
              ],
            },
          ],
        },
      ],
    } satisfies OdsWorkbookOptions;

    const parsed = await parseWorkbook(await generateWorkbook(workbook));
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.style).toMatchObject({
      font: { font: "Liberation Sans" },
      alignment: { textRotation: 45, shrinkToFit: true },
    });
  });

  it("projects cell hyperlinks into the canonical worksheet model", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles/><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p><text:a xlink:type="simple" xlink:href="https://example.com">Open</text:a></text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
      "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles/></office:document-styles>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.hyperlinks).toEqual([
      { cell: "A1", url: "https://example.com", display: "Open" },
    ]);

    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]?.hyperlinks).toEqual(parsed.worksheets?.[0]?.hyperlinks);
  });

  it("round-trips ODF table scenarios and linked sources", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles/><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-source table:mode="copy-results-only" table:table-name="Remote" xlink:type="simple" xlink:href="https://example.com/data.csv" xlink:actuate="onRequest" table:filter-name="Rows" table:filter-options="active" table:refresh-delay="PT1H"/><table:scenario table:scenario-ranges="Sheet1.A1:B2" table:is-active="true" table:display-border="false" table:border-color="#112233" table:copy-back="true" table:copy-styles="false" table:copy-formulas="true" table:comment="Scenario" table:protected="true"/><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.tableSource).toEqual({
      href: "https://example.com/data.csv",
      mode: "copy-results-only",
      tableName: "Remote",
      actuate: "onRequest",
      filterName: "Rows",
      filterOptions: "active",
      refreshDelay: "PT1H",
    });
    expect(parsed.worksheets?.[0]?.scenario).toEqual({
      ranges: "Sheet1.A1:B2",
      active: true,
      displayBorder: false,
      borderColor: "#112233",
      copyBack: true,
      copyStyles: false,
      copyFormulas: true,
      comment: "Scenario",
      protected: true,
    });
    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]?.tableSource).toEqual(parsed.worksheets?.[0]?.tableSource);
    expect(regenerated.worksheets?.[0]?.scenario).toEqual(parsed.worksheets?.[0]?.scenario);
  });

  it("round-trips ODF table title and description", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:title>Quarterly report</table:title><table:desc>Imported from finance</table:desc><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]).toMatchObject({
      tableTitle: "Quarterly report",
      tableDescription: "Imported from finance",
    });

    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]).toMatchObject({
      tableTitle: "Quarterly report",
      tableDescription: "Imported from finance",
    });
  });

  it("rejects invalid ODF table source attributes", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-source table:mode="invalid" xlink:type="simple" xlink:href="source.csv"/><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    await expect(parseWorkbook(source)).rejects.toMatchObject({
      name: "OdsParseError",
      elementName: "table:mode",
      reason: "invalid table source mode",
    });
  });

  it("inherits table-cell styles across common and automatic styles", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="Child" style:family="table-cell" style:parent-style-name="Parent"><style:table-cell-properties style:shrink-to-fit="true"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Child" office:value-type="string"><text:p>Styled</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
      "styles.xml": `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}><office:styles><style:style style:name="Parent" style:family="table-cell"><style:text-properties fo:font-family="Liberation Sans" fo:font-weight="bold"/></style:style></office:styles></office:document-styles>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.style).toMatchObject({
      font: { font: "Liberation Sans", bold: true },
      alignment: { shrinkToFit: true },
    });
  });

  it("rejects cyclic table-cell style inheritance", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="First" style:family="table-cell" style:parent-style-name="Second"/><style:style style:name="Second" style:family="table-cell" style:parent-style-name="First"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="First"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    });
    await expect(parseWorkbook(source)).rejects.toMatchObject({
      name: "OdsParseError",
      elementName: "style:parent-style-name",
      reason: "cyclic-style-inheritance",
    });
  });

  it("round-trips Gnumeric selection state through canonical selections", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="Table" style:family="table"><style:table-properties table:display="true" style:writing-mode="lr-tb" gnm:display-formulas="false" gnm:display-col-header="true" gnm:display-row-header="true"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1" table:style-name="Table"><table:table-row><table:table-cell/></table:table-row><gnm:selections gnm:cursor-col="8" gnm:cursor-row="176"><gnm:selection gnm:start-col="8" gnm:start-row="176" gnm:end-col="9" gnm:end-row="177"/></gnm:selections></table:table></office:spreadsheet></office:body></office:document-content>`;
    const parsed = await parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    );
    expect(parsed.worksheets?.[0]?.selection).toEqual([{ activeCell: "I177", sqref: "I177:J178" }]);
    expect(parsed.worksheets?.[0]?.sheetView).toEqual({
      showFormulas: false,
      showRowColHeaders: true,
    });
    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]?.selection).toEqual(parsed.worksheets?.[0]?.selection);
    expect(regenerated.worksheets?.[0]?.sheetView).toEqual(parsed.worksheets?.[0]?.sheetView);
  });

  it("round-trips merged cells through spans and covered cells", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2" office:value-type="string"><text:p>Merged</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const source = await parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    );
    expect(source.worksheets?.[0]?.mergeCells).toEqual([{ ref: "A1:B2" }]);

    const roundTrip = await parseWorkbook(await generateWorkbook(source));
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

  it("generates covered cells on every row of an authored merge", async () => {
    const generated = await generateWorkbook({
      worksheets: [
        {
          mergeCells: [{ ref: "A2:B3" }],
          rows: [{ rowNumber: 2, cells: [{ reference: "A2", value: "Merged" }] }],
        },
      ],
    });
    const parsed = await parseWorkbook(generated);
    expect(parsed.worksheets?.[0]?.mergeCells).toEqual([{ ref: "A2:B3" }]);
    const content = (await readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet"))
      .files["content.xml"]!;
    expect(content).toContain('<table:covered-table-cell table:number-columns-repeated="2"/>');
    expect(parsed.worksheets?.[0]?.rows?.[2]?.cells?.map((cell) => cell.reference)).toEqual([
      "A3",
      "B3",
    ]);
  });

  it("preserves row spans across repeated merged blocks", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2" office:value-type="string"><text:p>Merged</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="string"><text:p>After</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const parsed = await parseWorkbook(
      generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
    );
    expect(parsed.worksheets?.[0]?.mergeCells).toEqual([
      { ref: "A1:B2" },
      { ref: "A3:B4" },
      { ref: "A5:B6" },
    ]);
    expect(parsed.worksheets?.[0]?.rows?.map((row) => row.rowNumber)).toEqual([1, 3, 5, 7]);
  });

  it("requires an embedded object body but tolerates a missing replacement", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell><draw:frame svg:width="1cm" svg:height="1cm"><draw:object-ole xlink:href="./Object 1" xlink:type="simple"/></draw:frame></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": content,
      "Object 1": new Uint8Array([1]),
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.packageMembers).toEqual([
      { path: "Object 1", mediaType: "application/binary", data: new Uint8Array([1]) },
    ]);
    expect((await parseWorkbook(await generateWorkbook(parsed))).packageMembers).toEqual(
      parsed.packageMembers,
    );

    let error: unknown;
    try {
      await parseWorkbook(
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

  it("round-trips real chart anchors, semantics, worksheet links, and order", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
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

  it("assigns unique default chart subdocument paths across worksheets", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [
          {
            name: "First",
            charts: [{ type: "bar", series: [{ values: [1] }], col: 1, row: 1 }],
          },
          {
            name: "Second",
            charts: [{ type: "line", series: [{ values: [2] }], col: 1, row: 1 }],
          },
        ],
      }),
    );
    expect(parsed.worksheets?.map((worksheet) => worksheet.charts?.[0]?.name)).toEqual([
      "Object 1",
      "Object 2",
    ]);
  });

  it("preserves chart frame names separately from chart subdocument paths", async () => {
    const generated = await generateWorkbook({
      worksheets: [
        {
          name: "Sheet1",
          charts: [
            {
              name: "Visible chart",
              sourcePath: "Charts/Original chart",
              type: "line",
              series: [{ values: [1] }],
              col: 1,
              row: 1,
            },
          ],
        },
      ],
    });
    const parsed = await parseWorkbook(generated);
    expect(parsed.worksheets?.[0]?.charts?.[0]).toMatchObject({
      name: "Visible chart",
      sourcePath: "Charts/Original chart",
    });
    const { files } = await readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet");
    expect(files["Charts/Original chart/content.xml"]).toBeDefined();
    expect(files["content.xml"]).toContain('draw:name="Visible chart"');
    expect(files["content.xml"]).toContain('xlink:href="Charts/Original chart"');
  });

  it("rejects duplicate explicit chart subdocument paths", async () => {
    await expect(
      generateWorkbook({
        worksheets: [
          {
            name: "First",
            charts: [{ name: "Shared", type: "bar", series: [], col: 1, row: 1 }],
          },
          {
            name: "Second",
            charts: [{ name: "Shared", type: "line", series: [], col: 1, row: 1 }],
          },
        ],
      }),
    ).rejects.toMatchObject({
      name: "OdsGenerateError",
      reason: "duplicate-chart-path",
    });
  });

  it("round-trips defined names and rejects unknown worksheet children", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
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
      await parseWorkbook(
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
      elementName: "table:unknown",
      reason: "element has no canonical WorksheetOptions mapping",
    });
  });

  it("emits linked-only cell images without a package media part", async () => {
    const image = {
      type: "image" as const,
      href: "https://example.com/logo.png",
      sourceUrl: "https://example.com/logo.png",
      width: 914400,
      height: 914400,
    };
    const generated = await generateWorkbook({
      worksheets: [{ rows: [{ cells: [{ value: "Logo", graphics: [image] }] }] }],
    });
    const { binaries, manifest } = await readOcf(
      generated,
      "application/vnd.oasis.opendocument.spreadsheet",
    );

    expect(Object.keys(binaries)).toEqual([]);
    expect(manifest.entries.map((entry) => entry.fullPath)).not.toContain(image.href);
    expect(
      (await parseWorkbook(generated)).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics?.[0],
    ).toMatchObject({ sourceUrl: image.sourceUrl, data: undefined });
  });

  it("restores percent and length semantics in style overlays", async () => {
    const options: OdsWorkbookOptions = {
      worksheets: [{ rows: [{ cells: [{ value: "Overlay" }] }] }],
      styleResources: [
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
              attributes: {
                "fo:line-height": 150,
                "fo:margin-left": 38,
                "style:contextual-spacing": true,
              },
              children: [
                {
                  name: "style:tab-stops",
                  stops: [{ attributes: { "style:position": 914400, "style:type": "right" } }],
                },
              ],
            },
          ],
          maps: [
            {
              name: "style:map",
              condition: "cell-content()=1",
              applyStyleName: "Strong",
              baseCellAddress: "Sheet1.A1",
            },
          ],
        },
      ],
    };
    const generated = await generateWorkbook(options);
    const content = (await readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet"))
      .files["content.xml"]!;
    expect(content).toContain('style:row-height="20pt"');
    expect(content).toContain('fo:line-height="150%"');
    expect(content).toContain('fo:margin-left="1.005417cm"');
    expect(content).toContain('style:contextual-spacing="true"');
    expect(content).not.toMatch(/(?:row-height|line-height)="(?:150|20)"(?:\s|\/>)/);
    expect(await parseWorkbook(generated)).toMatchObject({
      styleResources: [
        {
          name: "P1",
          family: "paragraph",
          properties: [
            {
              name: "style:paragraph-properties",
              attributes: {
                "fo:line-height": 150,
                "fo:margin-left": 38,
                "style:contextual-spacing": true,
              },
              children: [
                {
                  name: "style:tab-stops",
                  stops: [{ attributes: { "style:position": 914400, "style:type": "right" } }],
                },
              ],
            },
          ],
          maps: options.styleResources?.[1]?.maps,
        },
      ],
    });
  });

  it("rejects a cell image with neither embedded data nor a source URL", async () => {
    let error: unknown;
    try {
      await generateWorkbook({
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
      elementName: "draw:image",
      reason: "missing-image-source",
    });
  });

  it("treats parent-relative image references as external source URLs", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value="1"><draw:frame svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><draw:image xlink:href="../logo.png"/><svg:title>Logo</svg:title></draw:frame></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const graphic = (
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", { "content.xml": content }),
      )
    ).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics?.[0];
    expect(graphic).toMatchObject({
      type: "image",
      href: "../logo.png",
      sourceUrl: "../logo.png",
      title: "Logo",
    });
  });

  it("round-trips cell object frames and rejects unknown frame children", async () => {
    const object = {
      type: "object" as const,
      href: "embedded-object",
      name: "Legacy object",
      description: "Embedded legacy object",
      notifyOnUpdateOfRanges: "A1:B2 C3:D4",
      x: 360000,
      y: 720000,
      width: 2160000,
      height: 1440000,
    };
    const generated = await generateWorkbook({
      worksheets: [{ rows: [{ cells: [{ value: 1, graphics: [object] }] }] }],
    });
    expect(
      (await readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet")).files[
        "content.xml"
      ],
    ).toContain('draw:notify-on-update-of-ranges="A1:B2 C3:D4"');
    const parsed = await parseWorkbook(generated);
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics).toEqual([object]);

    const objectContent = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value="1"><draw:frame draw:name="Legacy object" svg:x="1cm" svg:y="2cm" svg:width="6cm" svg:height="4cm"><draw:object xlink:href="./embedded-object"/></draw:frame></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    expect(
      (
        await parseWorkbook(
          generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
            "content.xml": objectContent,
          }),
        )
      ).worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics,
    ).toMatchObject([
      {
        type: "object",
        href: "embedded-object",
        name: "Legacy object",
        x: 360000,
        y: 720000,
        width: 2160000,
        height: 1440000,
      },
    ]);

    const oleContent = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value="1"><draw:frame draw:name="Legacy object" svg:x="1cm" svg:y="2cm" svg:width="6cm" svg:height="4cm"><draw:object-ole xlink:href="./embedded-object" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><svg:desc>Embedded legacy object</svg:desc><draw:unknown/></draw:frame></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    let error: unknown;
    try {
      await parseWorkbook(
        generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
          "content.xml": oleContent,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdsParseError);
    expect(error).toMatchObject({ part: "content.xml", elementName: "draw:unknown" });
  });

  it("round-trips calculation settings, forms, and covered cells", async () => {
    const calculation = {
      iterate: true,
      iterateCount: 5,
      iterateDelta: 0.01,
      caseSensitive: false,
      precisionAsShown: true,
      automaticLabelSearch: true,
      regularExpressionSearch: true,
      wildcardSearch: true,
      searchCriteriaMustApplyToWholeCell: true,
      nullYear: 1930,
      nullDate: { year: 1900, month: 1, day: 1 },
    };
    const controls: FormControlOptions[] = [
      { control: "text", id: "text", xmlId: "text", value: "", currentValue: "", maxLength: 0 },
      { control: "textArea", id: "textArea", xmlId: "textArea", paragraphs: ["Line"] },
      {
        control: "formattedText",
        id: "formattedText",
        xmlId: "formattedText",
        minValue: "0",
        maxValue: "9",
        validation: false,
      },
      {
        control: "number",
        id: "number",
        xmlId: "number",
        value: 0,
        currentValue: 1,
        minValue: 0,
        maxValue: 9,
      },
      { control: "date", id: "date", xmlId: "date", value: "2026-01-31" },
      { control: "time", id: "time", xmlId: "time", value: "01:02:03" },
      { control: "comboBox", id: "comboBox", xmlId: "comboBox", items: [{ label: "" }] },
      {
        control: "listBox",
        id: "listBox",
        xmlId: "listBox",
        options: [{ label: "", currentSelected: false, selected: false, value: "" }],
      },
      {
        control: "checkbox",
        id: "checkbox",
        xmlId: "checkbox",
        currentState: "unknown",
        state: "unknown",
        triState: true,
      },
      { control: "password", id: "password", xmlId: "password", echoChar: "*" },
      { control: "file", id: "file", xmlId: "file", currentValue: "a.ods" },
      { control: "fixedText", id: "fixedText", xmlId: "fixedText", label: "", multiLine: true },
      {
        control: "button",
        id: "button",
        xmlId: "button",
        buttonType: "push",
        defaultButton: false,
        toggle: false,
        focusOnClick: false,
      },
      { control: "image", id: "image", xmlId: "image", buttonType: "url", href: "logo.png" },
      { control: "radio", id: "radio", xmlId: "radio", currentSelected: false, selected: false },
      { control: "frame", id: "frame", xmlId: "frame", label: "Frame" },
      { control: "imageFrame", id: "imageFrame", xmlId: "imageFrame", imageData: "logo.png" },
      { control: "hidden", id: "hidden", xmlId: "hidden", value: "" },
      {
        control: "grid",
        id: "grid",
        xmlId: "grid",
        columns: [{ name: "title", controls: [{ control: "text", id: "title", xmlId: "title" }] }],
      },
      {
        control: "valueRange",
        id: "valueRange",
        xmlId: "valueRange",
        minValue: 0,
        maxValue: 100,
        stepSize: 1,
        pageStepSize: 10,
      },
      { control: "genericControl", id: "genericControl", xmlId: "genericControl" },
    ];
    const forms: OfficeFormsOptions = {
      automaticFocus: true,
      designMode: true,
      containers: [
        {
          name: "Inputs",
          controls,
        },
      ],
    };
    const parsed = await parseWorkbook(
      await generateWorkbook({
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

  it("preserves an iteration element with omitted status", async () => {
    const source = generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
      "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:calculation-settings><table:iteration table:maximum-difference="0.0001"/></table:calculation-settings></office:spreadsheet></office:body></office:document-content>`,
    });
    const parsed = await parseWorkbook(source);
    expect(parsed.calculation).toMatchObject({ iterationPresent: true, iterateDelta: 0.0001 });
    const content = (
      await readOcf(
        await generateWorkbook(parsed),
        "application/vnd.oasis.opendocument.spreadsheet",
      )
    ).files["content.xml"]!;
    expect(content).toContain('<table:iteration table:maximum-difference="0.0001"/>');
  });

  it("round-trips canonical spreadsheet error values", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [{ rows: [{ cells: [{ reference: "A1", error: "#DIV/0!" }] }] }],
      }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.error).toBe("#DIV/0!");
  });

  it("keeps error cells typed as errors when a number format is present", async () => {
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [
          {
            rows: [
              {
                cells: [
                  {
                    reference: "A1",
                    error: "#VALUE!",
                    style: { numFmt: "$#,##0.00" },
                  },
                ],
              },
            ],
          },
        ],
      }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]).toMatchObject({
      error: "#VALUE!",
    });
  });

  it("rejects malformed office forms with structured ODS errors", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:table><office:forms form:automatic-focus="yes"/></table:table></office:spreadsheet></office:body></office:document-content>`;
    let error: unknown;
    try {
      await parseWorkbook(
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
      elementName: "form:automatic-focus",
      reason: "invalid boolean value: yes",
    });
  });

  it("round-trips ODF form-control shapes with typed references", async () => {
    const shape = {
      col: 1,
      row: 1,
      controlId: "control1",
      name: "Control",
      properties: { geometry: "rect" },
    } satisfies NonNullable<WorksheetOptions["shapes"]>[number];
    const parsed = await parseWorkbook(
      await generateWorkbook({
        worksheets: [{ name: "Sheet1", rows: [{ cells: [{ reference: "A1" }] }], shapes: [shape] }],
      }),
    );
    expect(parsed.worksheets?.[0]?.shapes?.[0]).toMatchObject(shape);
    const regenerated = await parseWorkbook(await generateWorkbook(parsed));
    expect(regenerated.worksheets?.[0]?.shapes?.[0]).toEqual(parsed.worksheets?.[0]?.shapes?.[0]);
  });

  it("rejects unknown calculation-setting children", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:spreadsheet><table:calculation-settings><table:unknown/></table:calculation-settings></office:spreadsheet></office:body></office:document-content>`;
    let error: unknown;
    try {
      await parseWorkbook(
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
      elementName: "table:unknown",
      reason: "element has no canonical calculation mapping",
    });
  });
});
