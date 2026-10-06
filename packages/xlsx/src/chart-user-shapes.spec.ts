/**
 * Chart user-shapes wiring: fresh authoring emits the companion part, the
 * chart part's own rels entry, and the content-type Override; parse reads
 * the body back through that rels entry.
 *
 * @module
 */
import { unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./index";
import { parseWorkbookSync } from "./parse";
import type { WorkbookOptions } from "./parts/file";

const WORKBOOK: WorkbookOptions = {
  worksheets: [
    {
      name: "Data",
      rows: [{ cells: [{ value: "A" }, { value: 1 }] }],
      charts: [
        {
          type: "column",
          title: "Annotated",
          categories: ["A"],
          series: [{ name: "S", values: [1] }],
          col: 4,
          row: 1,
          userShapes: {
            anchors: [
              {
                from: { x: 0.1, y: 0.1 },
                to: { x: 0.4, y: 0.3 },
                object: {
                  type: "shape",
                  id: 1,
                  shapeProperties: { geometry: "rect", fill: { type: "solid", color: "FF0000" } },
                  style: {
                    lineReference: { index: 2 },
                    fillReference: { index: 1 },
                    effectReference: { index: 0 },
                    fontReference: { collection: "minor" },
                  },
                },
              },
            ],
          },
        },
      ],
    },
  ],
};

function fileText(entries: Record<string, Uint8Array>, name: string): string {
  const data = entries[name];
  if (!data) throw new Error(`missing part: ${name}`);
  return new TextDecoder().decode(data);
}

function injectChartExternalLink(archive: Record<string, Uint8Array>): void {
  archive["xl/charts/chart1.xml"] = new TextEncoder().encode(
    new TextDecoder()
      .decode(archive["xl/charts/chart1.xml"]!)
      .replace(
        "</c:chartSpace>",
        '<c:externalData r:id="rId7"><c:autoUpdate val="1"/></c:externalData></c:chartSpace>',
      ),
  );
  archive["xl/charts/_rels/chart1.xml.rels"] = new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../externalLinks/externalLink7.xml"/>` +
      `</Relationships>`,
  );
  archive["xl/externalLinks/externalLink7.xml"] = new TextEncoder().encode(
    `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<externalBook r:id="rId3"><sheetNames><sheetName val="Source"/></sheetNames></externalBook></externalLink>`,
  );
  archive["xl/externalLinks/_rels/externalLink7.xml.rels"] = new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="https://example.com/source.xlsx" TargetMode="External"/>` +
      `</Relationships>`,
  );
  archive["[Content_Types].xml"] = new TextEncoder().encode(
    new TextDecoder()
      .decode(archive["[Content_Types].xml"]!)
      .replace(
        "</Types>",
        `<Override PartName="/xl/externalLinks/externalLink7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>`,
      ),
  );
}

describe("chart userShapes companion part", () => {
  it("emits the part, chart rels entry, and content-type Override", async () => {
    const bytes = await generateWorkbook(WORKBOOK);
    const entries = unzipSync(new Uint8Array(bytes));

    const chart = fileText(entries, "xl/charts/chart1.xml");
    expect(chart).toContain('<c:userShapes r:id="rId1"/>');

    const shapes = fileText(entries, "xl/charts/userShapes1.xml");
    expect(shapes).toContain("<cdr:relSizeAnchor>");
    expect(shapes).toContain("<cdr:sp>");
    expect(shapes).toContain("<cdr:style>");
    expect(shapes).toContain('<a:prstGeom prst="rect">');

    const rels = fileText(entries, "xl/charts/_rels/chart1.xml.rels");
    expect(rels).toContain('Id="rId1"');
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes"',
    );
    expect(rels).toContain('Target="userShapes1.xml"');

    const contentTypes = fileText(entries, "[Content_Types].xml");
    expect(contentTypes).toContain(
      'PartName="/xl/charts/userShapes1.xml"' +
        ' ContentType="application/vnd.openxmlformats-officedocument.drawingml.chartUserShapes+xml"',
    );
  });

  it("round-trips the anchors through the chart part rels", async () => {
    const bytes = await generateWorkbook(WORKBOOK);
    const parsed = parseWorkbookSync(bytes);
    const chart = parsed.worksheets?.[0]?.charts?.[0];
    expect(chart?.userShapes?.relationshipId).toBe("rId1");
    expect(chart?.userShapes?.anchors).toHaveLength(1);
    const anchor = chart?.userShapes?.anchors?.[0];
    if (!anchor || !("to" in anchor)) throw new Error("expected a relative anchor");
    expect(anchor.from).toEqual({ x: 0.1, y: 0.1 });
    expect(anchor.object.type).toBe("shape");
    if (anchor.object.type !== "shape") throw new Error("expected a shape object");
    expect(anchor.object.shapeProperties.geometry).toEqual({ preset: "rect" });
    expect(anchor.object.style).toEqual({
      lineReference: { index: 2 },
      fillReference: { index: 1 },
      effectReference: { index: 0 },
      fontReference: { collection: "minor" },
    });
  });

  it("round-trips a legacy companion root element", async () => {
    const bytes = (await generateWorkbook(WORKBOOK, { type: "uint8array" })) as Uint8Array;
    const archive = unzipSync(bytes);
    archive["xl/charts/userShapes1.xml"] = new TextEncoder().encode(
      fileText(archive, "xl/charts/userShapes1.xml")
        .replace("<cdr:userShapes ", "<c:userShapes ")
        .replace("</cdr:userShapes>", "</c:userShapes>"),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    const chart = parsed.worksheets?.[0]?.charts?.[0];
    expect(chart?.userShapes?.rootElement).toBe("chart");

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const result = unzipSync(output);
    const shapes = fileText(result, "xl/charts/userShapes1.xml");
    expect(shapes).toContain("<c:userShapes ");
    expect(shapes).toContain("</c:userShapes>");
  });

  it("preserves companion parts outside charts directory", async () => {
    const bytes = (await generateWorkbook(WORKBOOK, { type: "uint8array" })) as Uint8Array;
    const archive = unzipSync(bytes);
    archive["xl/drawings/chartShapes1.xml"] = archive["xl/charts/userShapes1.xml"]!;
    delete archive["xl/charts/userShapes1.xml"];
    archive["xl/charts/_rels/chart1.xml.rels"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["xl/charts/_rels/chart1.xml.rels"]!)
        .replace('Target="userShapes1.xml"', 'Target="../drawings/chartShapes1.xml"'),
    );
    archive["[Content_Types].xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["[Content_Types].xml"]!)
        .replace("/xl/charts/userShapes1.xml", "/xl/drawings/chartShapes1.xml"),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    const chart = parsed.worksheets?.[0]?.charts?.[0];
    expect(chart?.userShapes?.path).toBe("xl/drawings/chartShapes1.xml");

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const result = unzipSync(output);
    expect(result["xl/drawings/chartShapes1.xml"]).toBeDefined();
    expect(result["xl/charts/userShapes1.xml"]).toBeUndefined();
    expect(fileText(result, "xl/charts/_rels/chart1.xml.rels")).toContain(
      'Target="../drawings/chartShapes1.xml"',
    );
  });

  it("rewires chart external data to its typed external link part", async () => {
    const source = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Data",
            rows: [{ cells: [{ value: "A" }] }],
            charts: [
              {
                type: "column",
                categories: ["A"],
                series: [{ name: "S", values: [1] }],
                col: 4,
                row: 1,
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    archive["xl/charts/chart1.xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["xl/charts/chart1.xml"]!)
        .replace(
          "</c:chartSpace>",
          '<c:externalData r:id="rId7"><c:autoUpdate val="1"/></c:externalData></c:chartSpace>',
        ),
    );
    archive["xl/charts/_rels/chart1.xml.rels"] = new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../externalLinks/externalLink7.xml"/>` +
        `</Relationships>`,
    );
    archive["xl/externalLinks/externalLink7.xml"] = new TextEncoder().encode(
      `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<externalBook r:id="rId3"><sheetNames><sheetName val="Source"/></sheetNames></externalBook></externalLink>`,
    );
    archive["xl/externalLinks/_rels/externalLink7.xml.rels"] = new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="https://example.com/source.xlsx" TargetMode="External"/>` +
        `</Relationships>`,
    );
    archive["[Content_Types].xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["[Content_Types].xml"]!)
        .replace(
          "</Types>",
          `<Override PartName="/xl/externalLinks/externalLink7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>`,
        ),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    const chart = parsed.worksheets![0]!.charts![0]!;
    expect(chart.externalData).toMatchObject({ relationshipId: "rId7", autoUpdate: true });
    expect(chart.externalDataRelationshipTarget).toBe("../externalLinks/externalLink7.xml");
    expect(chart.externalLinkPath).toBe("xl/externalLinks/externalLink7.xml");
    expect(chart.externalLink?.externalBook).toMatchObject({
      sheetNames: ["Source"],
      target: "https://example.com/source.xlsx",
    });
    expect(parsed.rawParts).toBeUndefined();

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const result = unzipSync(output);
    expect(fileText(result, "xl/charts/chart1.xml")).toContain('r:id="rId7"');
    const chartRels = fileText(result, "xl/charts/_rels/chart1.xml.rels");
    expect(chartRels).toContain('Id="rId7"');
    expect(chartRels).toContain('Target="../externalLinks/externalLink7.xml"');
    expect(fileText(result, "xl/externalLinks/externalLink7.xml")).toContain('r:id="rId3"');
    expect(fileText(result, "xl/externalLinks/_rels/externalLink7.xml.rels")).toContain(
      'Target="https://example.com/source.xlsx"',
    );
  });

  it("rewires chartsheet chart external data to its typed external link part", async () => {
    const source = (await generateWorkbook(
      {
        chartsheets: [
          {
            name: "Chart",
            chart: { type: "column", series: [{ name: "S", values: [1] }] },
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    injectChartExternalLink(archive);

    const parsed = parseWorkbookSync(zipSync(archive));
    const sheet = parsed.chartsheets![0]!;
    expect(sheet.chart?.externalData).toMatchObject({ relationshipId: "rId7", autoUpdate: true });
    expect(sheet.externalLinkPath).toBe("xl/externalLinks/externalLink7.xml");
    expect(sheet.externalLink?.externalBook).toMatchObject({
      target: "https://example.com/source.xlsx",
    });

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const result = unzipSync(output);
    expect(fileText(result, "xl/charts/_rels/chart1.xml.rels")).toContain(
      'Target="../externalLinks/externalLink7.xml"',
    );
    expect(fileText(result, "xl/externalLinks/_rels/externalLink7.xml.rels")).toContain(
      'Target="https://example.com/source.xlsx"',
    );
  });
});
