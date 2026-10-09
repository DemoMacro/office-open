import { unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./index";
import { parseWorkbookSync } from "./parse";

function fileText(entries: Record<string, Uint8Array>, name: string): string {
  const data = entries[name];
  if (!data) throw new Error(`missing part: ${name}`);
  return new TextDecoder().decode(data);
}

function replaceText(entries: Record<string, Uint8Array>, name: string, text: string): void {
  entries[name] = new TextEncoder().encode(text);
}

const CHART = {
  type: "column",
  categories: ["A"],
  series: [{ name: "S", values: [1] }],
} as const;

describe("xlsx relationship topology", () => {
  it("preserves unmodeled drawing companion relationships", async () => {
    const source = (await generateWorkbook({
      worksheets: [
        {
          name: "Data",
          rows: [{ cells: [{ value: "A" }] }],
          images: [{ data: "AAAA", type: "jpg", col: 1, row: 1 }],
          charts: [{ ...CHART, col: 1, row: 1 }],
        },
      ],
    })) as Uint8Array;
    const archive = unzipSync(source);
    const rels = fileText(archive, "xl/drawings/_rels/drawing1.xml.rels").replace(
      "</Relationships>",
      '<Relationship Id="rId8" Type="https://example.com/companion" Target="../media/image1.jpeg"/></Relationships>',
    );
    replaceText(archive, "xl/drawings/_rels/drawing1.xml.rels", rels);

    const output = unzipSync(
      (await generateWorkbook(parseWorkbookSync(zipSync(archive)))) as Uint8Array,
    );
    const result = fileText(output, "xl/drawings/_rels/drawing1.xml.rels");
    expect(result).toContain('Id="rId8"');
    expect(result).toContain('Type="https://example.com/companion"');
    expect(result).toContain('Target="../media/image1.jpeg"');
  });

  it("preserves internal drawing hyperlink targets without TargetMode", async () => {
    const source = (await generateWorkbook({
      worksheets: [
        {
          name: "Data",
          rows: [{ cells: [{ value: "A" }] }],
          shapes: [
            {
              col: 1,
              row: 1,
              name: "Shape",
              properties: { geometry: "rect" },
              hyperlink: { url: "#Sheet2!A1" },
            },
          ],
        },
      ],
    })) as Uint8Array;
    const archive = unzipSync(source);
    replaceText(
      archive,
      "xl/drawings/_rels/drawing1.xml.rels",
      fileText(archive, "xl/drawings/_rels/drawing1.xml.rels").replace(
        ' TargetMode="External"',
        "",
      ),
    );

    const output = unzipSync(
      (await generateWorkbook(parseWorkbookSync(zipSync(archive)))) as Uint8Array,
    );
    const result = fileText(output, "xl/drawings/_rels/drawing1.xml.rels");
    expect(result).toContain('Target="#Sheet2!A1"');
    expect(result).not.toContain("TargetMode");
  });

  it("preserves chartsheet drawing and chart paths", async () => {
    const source = (await generateWorkbook({
      chartsheets: [
        { name: "One", chart: CHART },
        { name: "Two", chart: CHART },
      ],
    })) as Uint8Array;
    const archive = unzipSync(source);
    replaceText(
      archive,
      "xl/chartsheets/_rels/sheet1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/><Relationship Id="rId1" Type="https://example.com/companion" Target="../companion.xml"/></Relationships>`,
    );
    replaceText(archive, "xl/companion.xml", "<companion/>");
    replaceText(
      archive,
      "xl/chartsheets/_rels/sheet2.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    );
    replaceText(
      archive,
      "xl/drawings/_rels/drawing1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>`,
    );
    replaceText(
      archive,
      "xl/drawings/_rels/drawing2.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.chartsheets?.[0]?.sourceDrawingPath).toBe("xl/drawings/drawing2.xml");
    expect(parsed.chartsheets?.[0]?.sourceChartPath).toBe("xl/charts/chart1.xml");
    expect(parsed.chartsheets?.[0]?.sourceDrawingRelationshipId).toBe("rId3");
    expect(parsed.chartsheets?.[1]?.sourceDrawingPath).toBe("xl/drawings/drawing1.xml");
    expect(parsed.chartsheets?.[1]?.sourceChartPath).toBe("xl/charts/chart2.xml");

    const output = unzipSync((await generateWorkbook(parsed)) as Uint8Array);
    expect(fileText(output, "xl/chartsheets/sheet1.xml")).toContain('<drawing r:id="rId3"/>');
    expect(fileText(output, "xl/chartsheets/_rels/sheet1.xml.rels")).toContain('Id="rId3"');
    expect(fileText(output, "xl/chartsheets/_rels/sheet1.xml.rels")).toContain(
      'Target="../drawings/drawing2.xml"',
    );
    expect(fileText(output, "xl/chartsheets/sheet1.xml")).toContain('<drawing r:id="rId3"/>');
    expect(fileText(output, "xl/chartsheets/_rels/sheet1.xml.rels")).toContain(
      '<Relationship Id="rId4" Type="https://example.com/companion" Target="../companion.xml"/>',
    );
    expect(fileText(output, "xl/chartsheets/_rels/sheet2.xml.rels")).toContain('Id="rId1"');
    expect(fileText(output, "xl/chartsheets/_rels/sheet2.xml.rels")).toContain(
      'Target="../drawings/drawing1.xml"',
    );
    expect(fileText(output, "xl/drawings/_rels/drawing1.xml.rels")).toContain(
      'Target="../charts/chart2.xml"',
    );
    expect(fileText(output, "xl/drawings/_rels/drawing2.xml.rels")).toContain(
      'Target="../charts/chart1.xml"',
    );
  });

  it("preserves duplicate external data consolidation relationships", async () => {
    const source = (await generateWorkbook({
      worksheets: [{ name: "Data", rows: [{ cells: [{ value: "A" }] }] }],
    })) as Uint8Array;
    const archive = unzipSync(source);
    replaceText(
      archive,
      "xl/worksheets/sheet1.xml",
      fileText(archive, "xl/worksheets/sheet1.xml").replace(
        "</worksheet>",
        '<dataConsolidate function="sum"><dataRefs count="2">' +
          '<dataRef ref="A1:B2" sheet="First" r:id="rId1"/>' +
          '<dataRef ref="C1:D2" sheet="Second" r:id="rId2"/>' +
          "</dataRefs></dataConsolidate></worksheet>",
      ),
    );
    const relationshipType =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
    replaceText(
      archive,
      "xl/worksheets/_rels/sheet1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationshipType}" Target="../externalLinks/source.xml" TargetMode="External"/><Relationship Id="rId2" Type="${relationshipType}" Target="../externalLinks/source.xml" TargetMode="External"/></Relationships>`,
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.worksheets?.[0]?.dataConsolidate?.refs).toEqual([
      { ref: "A1:B2", sheet: "First", rId: "rId1", target: "../externalLinks/source.xml" },
      { ref: "C1:D2", sheet: "Second", rId: "rId2", target: "../externalLinks/source.xml" },
    ]);
    const output = unzipSync(
      (await generateWorkbook(parseWorkbookSync(zipSync(archive)))) as Uint8Array,
    );
    const sheetXml = fileText(output, "xl/worksheets/sheet1.xml");
    const rels = fileText(output, "xl/worksheets/_rels/sheet1.xml.rels");
    expect(sheetXml).toContain('<dataRef ref="A1:B2" sheet="First" r:id="rId1"/>');
    expect(sheetXml).toContain('<dataRef ref="C1:D2" sheet="Second" r:id="rId2"/>');
    expect(rels.match(/Id="rId1"/)).toHaveLength(1);
    expect(rels.match(/Id="rId2"/)).toHaveLength(1);
    expect(rels.match(/Target="..\/externalLinks\/source.xml"/g)).toHaveLength(2);
  });

  it("preserves duplicate external hyperlink relationships", async () => {
    const source = (await generateWorkbook({
      worksheets: [{ name: "Data", rows: [{ cells: [{ value: "A" }] }] }],
    })) as Uint8Array;
    const archive = unzipSync(source);
    replaceText(
      archive,
      "xl/worksheets/sheet1.xml",
      fileText(archive, "xl/worksheets/sheet1.xml").replace(
        "</worksheet>",
        '<hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="B1" r:id="rId2"/></hyperlinks></worksheet>',
      ),
    );
    const relationshipType =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    replaceText(
      archive,
      "xl/worksheets/_rels/sheet1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationshipType}" Target="file:///Book2.xlsx" TargetMode="External"/><Relationship Id="rId2" Type="${relationshipType}" Target="file:///Book2.xlsx" TargetMode="External"/></Relationships>`,
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.worksheets?.[0]?.hyperlinks).toEqual([
      expect.objectContaining({ cell: "A1", relationshipId: "rId1" }),
      expect.objectContaining({ cell: "B1", relationshipId: "rId2" }),
    ]);
    const output = unzipSync(
      (await generateWorkbook(parseWorkbookSync(zipSync(archive)))) as Uint8Array,
    );
    const sheetXml = fileText(output, "xl/worksheets/sheet1.xml");
    const rels = fileText(output, "xl/worksheets/_rels/sheet1.xml.rels");
    expect(sheetXml).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(sheetXml).toContain('<hyperlink ref="B1" r:id="rId2"/>');
    expect(rels.match(/Id="rId1"/)).toHaveLength(1);
    expect(rels.match(/Id="rId2"/)).toHaveLength(1);
    expect(rels.match(/Target="file:\/\/\/Book2.xlsx"/g)).toHaveLength(2);
  });

  it("preserves a repair-style dangling calc chain relationship", async () => {
    const source = (await generateWorkbook({
      worksheets: [{ name: "Data", rows: [{ cells: [{ value: "A" }] }] }],
    })) as Uint8Array;
    const archive = unzipSync(source);
    delete archive["xl/calcChain.xml"];
    replaceText(
      archive,
      "xl/_rels/workbook.xml.rels",
      fileText(archive, "xl/_rels/workbook.xml.rels").replace(
        "</Relationships>",
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
      ),
    );

    const output = unzipSync(
      (await generateWorkbook(parseWorkbookSync(zipSync(archive)))) as Uint8Array,
    );
    expect(output["xl/calcChain.xml"]).toBeUndefined();
    const rels = fileText(output, "xl/_rels/workbook.xml.rels");
    expect(rels).toContain('Id="rId9"');
    expect(rels).toContain("/calcChain");
  });

  it("separates chartsheet names from chart frame names", async () => {
    const source = (await generateWorkbook({
      chartsheets: [
        { name: "Outdoor", chart: CHART },
        { name: "Logarithmic", chart: CHART },
      ],
    })) as Uint8Array;
    const archive = unzipSync(source);
    replaceText(
      archive,
      "xl/drawings/drawing1.xml",
      fileText(archive, "xl/drawings/drawing1.xml").replace('name="Chart 1"', 'name="Frame One"'),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.chartsheets?.map((sheet) => sheet.name)).toEqual(["Outdoor", "Logarithmic"]);
    expect(parsed.chartsheets?.[0]?.chartName).toBe("Frame One");

    const output = unzipSync((await generateWorkbook(parsed)) as Uint8Array);
    expect(fileText(output, "xl/workbook.xml")).toContain('name="Outdoor"');
    expect(fileText(output, "xl/drawings/drawing1.xml")).toContain('name="Frame One"');
  });

  it("preserves chart style and color companion parts", async () => {
    const source = (await generateWorkbook({
      worksheets: [{ name: "Data", charts: [{ ...CHART, col: 1, row: 1 }] }],
    })) as Uint8Array;
    const archive = unzipSync(source);
    archive["xl/charts/style1.xml"] = new TextEncoder().encode("<chartStyle/>");
    archive["xl/charts/colors1.xml"] = new TextEncoder().encode("<chartColorStyle/>");
    replaceText(
      archive,
      "xl/charts/_rels/chart1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>`,
    );
    archive["[Content_Types].xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["[Content_Types].xml"]!)
        .replace(
          "</Types>",
          `<Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>`,
        ),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.rawParts?.map((part) => part.path)).toEqual(
      expect.arrayContaining(["xl/charts/style1.xml", "xl/charts/colors1.xml"]),
    );

    const output = unzipSync((await generateWorkbook(parsed)) as Uint8Array);
    expect(output["xl/charts/style1.xml"]).toBeDefined();
    expect(output["xl/charts/colors1.xml"]).toBeDefined();
    const rels = fileText(output, "xl/charts/_rels/chart1.xml.rels");
    expect(rels).toContain("relationships/chartStyle");
    expect(rels).toContain('Target="style1.xml"');
    expect(rels).toContain("relationships/chartColorStyle");
    expect(rels).toContain('Target="colors1.xml"');
  });
});
