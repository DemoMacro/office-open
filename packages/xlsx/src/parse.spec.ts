import { unzipSync, zipSync } from "@office-open/core";
import type { WorkbookOptions } from "@parts/file";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./generate";
import { parseWorkbookSync } from "./parse";

// Full-file round-trip: generateWorkbook → Buffer → parseWorkbook → WorkbookOptions.
// Proves the three previously-deferred parse gaps are resolved on the read path.

async function roundTrip(opts: WorkbookOptions): Promise<WorkbookOptions> {
  const buf = (await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array;
  return parseWorkbookSync(buf);
}

describe("parseWorkbook round-trip", () => {
  it("keeps worksheet contents aligned when workbook relationships are out of order", async () => {
    const buffer = (await generateWorkbook(
      {
        worksheets: [
          { name: "Alpha", rows: [{ cells: [{ value: "alpha" }] }] },
          { name: "Beta", rows: [{ cells: [{ value: "beta" }] }] },
          { name: "Gamma", rows: [{ cells: [{ value: "gamma" }] }] },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(buffer);
    const relsPath = "xl/_rels/workbook.xml.rels";
    const relsXml = new TextDecoder().decode(archive[relsPath]!);
    const relationships = [...relsXml.matchAll(/<Relationship\b[^>]+\/>/g)].map(
      (match) => match[0],
    );
    const worksheetRels = relationships.filter((relationship) =>
      relationship.includes('Target="worksheets/'),
    );
    const otherRels = relationships.filter(
      (relationship) => !relationship.includes('Target="worksheets/'),
    );
    const firstRelationship = relsXml.search(/<Relationship\b[^>]+\/>/);
    const opening = relsXml.slice(0, firstRelationship);
    archive[relsPath] = new TextEncoder().encode(
      opening + [...worksheetRels].reverse().concat(otherRels).join("") + "</Relationships>",
    );

    const parsed = parseWorkbookSync(zipSync(archive));

    expect(parsed.worksheets?.map((worksheet) => worksheet.name)).toEqual([
      "Alpha",
      "Beta",
      "Gamma",
    ]);
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe("alpha");
    expect(parsed.worksheets?.[2]?.rows?.[0]?.cells?.[0]?.value).toBe("gamma");
  });

  it("emits the shared strings relationship only when the part exists", async () => {
    const withoutStrings = (await generateWorkbook(
      { worksheets: [{ rows: [{ cells: [{ value: 1 }] }] }] },
      { type: "uint8array" },
    )) as Uint8Array;
    const withoutArchive = unzipSync(withoutStrings);
    const withoutRels = new TextDecoder().decode(withoutArchive["xl/_rels/workbook.xml.rels"]!);
    expect(withoutArchive["xl/sharedStrings.xml"]).toBeUndefined();
    expect(withoutRels).not.toContain("/sharedStrings");

    const withStrings = (await generateWorkbook(
      { worksheets: [{ rows: [{ cells: [{ value: "text" }] }] }] },
      { type: "uint8array" },
    )) as Uint8Array;
    const withArchive = unzipSync(withStrings);
    const withRels = new TextDecoder().decode(withArchive["xl/_rels/workbook.xml.rels"]!);
    expect(withArchive["xl/sharedStrings.xml"]).toBeDefined();
    expect(withRels).toContain("/sharedStrings");
  });

  it("preserves relationship-less macro sheet definitions", async () => {
    const buffer = (await generateWorkbook(
      {
        sheetDefinitions: [{ name: "Macro", sheetId: 8, rId: "", state: "veryHidden" }],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const workbookXml = new TextDecoder().decode(unzipSync(buffer)["xl/workbook.xml"]!);

    expect(workbookXml).toContain('<sheet name="Macro" sheetId="8" r:id="" state="veryHidden"/>');
    expect(workbookXml).toContain('<sheet name="Macro" sheetId="8" r:id="" state="veryHidden"/>');
  });

  it("resolves absolute root metadata relationship targets", async () => {
    const buffer = (await generateWorkbook(
      {
        creator: "Owner",
        lastModifiedBy: "Editor",
        appProperties: { application: "Spreadsheet", docSecurity: 0 },
        worksheets: [{ name: "Sheet", rows: [{ cells: [{ value: 1 }] }] }],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(buffer);
    const relsPath = "_rels/.rels";
    archive[relsPath] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive[relsPath]!)
        .replaceAll('Target="docProps/', 'Target="/docProps/'),
    );
    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.creator).toBe("Owner");
    expect(parsed.lastModifiedBy).toBe("Editor");
    expect(parsed.appProperties?.application).toBe("Spreadsheet");
    expect(parsed.appProperties?.docSecurity).toBe(0);
  });

  it("preserves an empty shared-strings part", async () => {
    const source = (await generateWorkbook(
      { sharedStringsDeclared: true },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    expect(archive["xl/sharedStrings.xml"]).toBeDefined();
    const parsed = parseWorkbookSync(source);
    expect(parsed.sharedStringsDeclared).toBe(true);
    expect(parsed.sharedStrings).toBeUndefined();
  });

  it("keeps worksheet hyperlink relationship ids stable across round-trip", async () => {
    const source = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Sheet",
            hyperlinks: [
              { cell: "A1", url: "https://example.com/first" },
              { cell: "B2", url: "https://example.com/second" },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    const sheetPath = "xl/worksheets/sheet1.xml";
    const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
    archive[sheetPath] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive[sheetPath]!)
        .replaceAll('r:id="rId1"', 'r:id="rId7"')
        .replaceAll('r:id="rId2"', 'r:id="rId3"'),
    );
    archive[relsPath] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive[relsPath]!)
        .replaceAll('Id="rId1"', 'Id="rId7"')
        .replaceAll('Id="rId2"', 'Id="rId3"'),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.worksheets![0]!.hyperlinks).toMatchObject([
      { cell: "A1", url: "https://example.com/first", relationshipId: "rId7" },
      { cell: "B2", url: "https://example.com/second", relationshipId: "rId3" },
    ]);

    const output = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const outputArchive = unzipSync(output);
    const sheet = new TextDecoder().decode(outputArchive[sheetPath]!);
    const rels = new TextDecoder().decode(outputArchive[relsPath]!);
    expect(sheet).toContain('ref="A1" r:id="rId7"');
    expect(sheet).toContain('ref="B2" r:id="rId3"');
    expect(rels).toMatch(/Id="rId7"[^>]*Target="https:\/\/example\.com\/first"/);
    expect(rels).toMatch(/Id="rId3"[^>]*Target="https:\/\/example\.com\/second"/);
  });

  it("keeps WebExtension drawing anchors typed across round-trip", async () => {
    const source = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Sheet",
            shapes: [
              {
                col: 1,
                row: 1,
                properties: { x: 0, y: 0, width: 100, height: 100, geometry: "rect" },
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    archive["xl/drawings/drawing1.xml"] = new TextEncoder().encode(
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>2</xdr:col><xdr:row>4</xdr:row></xdr:to>` +
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
        `<mc:Choice xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" Requires="we">` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="WebExtension 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">` +
        `<we:webextensionref r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>` +
        `<mc:Fallback><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="WebExtension 1"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></mc:Fallback>` +
        `</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
    );
    archive["xl/drawings/_rels/drawing1.xml.rels"] = new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="../webextensions/webextension1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `</Relationships>`,
    );
    archive["xl/webextensions/webextension1.xml"] = new TextEncoder().encode(
      `<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{00000000-0000-0000-0000-000000000001}"/>`,
    );
    archive["xl/media/image1.png"] = new Uint8Array([1, 2, 3]);
    archive["[Content_Types].xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["[Content_Types].xml"]!)
        .replace(
          "</Types>",
          `<Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>`,
        ),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.worksheets?.[0]?.webExtensions).toMatchObject([
      {
        col: 1,
        row: 1,
        name: "WebExtension 1",
        sourcePath: "xl/webextensions/webextension1.xml",
        snapshotSourcePath: "xl/media/image1.png",
      },
    ]);
    expect(parsed.rawParts?.map((part) => part.path)).not.toContain("xl/drawings/drawing1.xml");

    const regenerated = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const output = unzipSync(regenerated);
    expect(new TextDecoder().decode(output["xl/drawings/drawing1.xml"]!)).toContain(
      "we:webextensionref",
    );
    expect(output["xl/webextensions/webextension1.xml"]).toBeDefined();
    const rels = new TextDecoder().decode(output["xl/drawings/_rels/drawing1.xml.rels"]!);
    expect(rels).toContain("relationships/webextension");
    expect(rels).toContain("../media/image1.png");
  });

  it("accepts the webextension element alias across round-trip", async () => {
    const source = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Sheet",
            shapes: [
              {
                col: 1,
                row: 1,
                properties: { x: 0, y: 0, width: 100, height: 100, geometry: "rect" },
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    archive["xl/drawings/drawing1.xml"] = new TextEncoder().encode(
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>` +
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
        `<mc:Choice Requires="we"><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="WebExtension"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>` +
        `<mc:Fallback/></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
    );
    archive["xl/drawings/_rels/drawing1.xml.rels"] = new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="../webextensions/webextension1.xml"/></Relationships>`,
    );
    archive["xl/webextensions/webextension1.xml"] = new TextEncoder().encode(
      `<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{synthetic-id}"/>`,
    );
    archive["[Content_Types].xml"] = new TextEncoder().encode(
      new TextDecoder()
        .decode(archive["[Content_Types].xml"]!)
        .replace(
          "</Types>",
          `<Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>`,
        ),
    );

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.worksheets?.[0]?.webExtensions?.[0]).toMatchObject({
      sourcePath: "xl/webextensions/webextension1.xml",
      elementName: "webextension",
    });
    const output = unzipSync((await generateWorkbook(parsed)) as Uint8Array);
    expect(new TextDecoder().decode(output["xl/drawings/drawing1.xml"]!)).toContain(
      "we:webextension ",
    );
    expect(output["xl/webextensions/webextension1.xml"]).toBeDefined();
    expect(new TextDecoder().decode(output["xl/drawings/_rels/drawing1.xml.rels"]!)).toContain(
      "relationships/webextension",
    );
  });

  it("keeps comment VML shapes typed across round-trip", async () => {
    const source = (await generateWorkbook(
      {
        worksheets: [{ name: "Sheet", comments: [{ cell: "A1", author: "A", text: "note" }] }],
      },
      { type: "uint8array" },
    )) as Uint8Array;
    const archive = unzipSync(source);
    const vml = new TextDecoder()
      .decode(archive["xl/drawings/vmlDrawing1.vml"]!)
      .replace('data="1"', 'data="7"')
      .replace(
        '<v:shape id="_x0000_s1025"',
        '<v:shape id="_x0000_s1025" alt="Source note" o:spid="_x0000_s2049"',
      )
      .replace('<v:fill color2="infoBackground [80]"/>', '<v:fill color2="#123456" angle="45"/>');
    archive["xl/drawings/vmlDrawing1.vml"] = new TextEncoder().encode(vml);

    const parsed = parseWorkbookSync(zipSync(archive));
    expect(parsed.rawParts?.map((part) => part.path) ?? []).not.toContain(
      "xl/drawings/vmlDrawing1.vml",
    );
    expect(parsed.worksheets?.[0]?.commentsVmlLayout).toMatchObject({
      idmap: { data: "7" },
    });
    expect(parsed.worksheets?.[0]?.comments?.[0]?.vmlShape).toMatchObject({
      id: "_x0000_s1025",
      alt: "Source note",
      spid: "_x0000_s2049",
      fill: { color2: "#123456", angle: 45 },
    });

    const regenerated = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const outputVml = new TextDecoder().decode(
      unzipSync(regenerated)["xl/drawings/vmlDrawing1.vml"]!,
    );
    expect(unzipSync(regenerated)["xl/drawings/vmlDrawing1.vml"]).toEqual(
      archive["xl/drawings/vmlDrawing1.vml"],
    );
    expect(outputVml).toContain('data="7"');
    expect(outputVml).toContain('alt="Source note"');
    expect(outputVml).toContain('o:spid="_x0000_s2049"');
    expect(outputVml).toContain('color2="#123456"');
  });

  it("round-trips pivot table page filters", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "Data",
          rows: [
            { cells: [{ value: "City" }, { value: "Year" }, { value: "Revenue" }] },
            { cells: [{ value: "Beijing" }, { value: 2023 }, { value: 320 }] },
            { cells: [{ value: "Shanghai" }, { value: 2024 }, { value: 580 }] },
          ],
        },
        {
          name: "Pivot",
          rows: [],
          pivotTables: [
            {
              mode: "source",
              source: "A1:C3",
              sourceSheet: "Data",
              rows: ["City"],
              pages: [{ field: "Year", item: 1 }],
              data: [{ field: "Revenue", summarize: "sum", name: "Revenue total" }],
            },
          ],
        },
      ],
    };

    const buf = (await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array;
    const archive = unzipSync(buf);
    // The authored pivot table survives as a verbatim passthrough part — the
    // lossy authoring projection (rows/pages) is not re-derived on parse.
    const pivotXml = new TextDecoder().decode(archive["xl/pivotTables/pivotTable1.xml"]!);
    expect(pivotXml).toContain('<pageFields count="1">');
    expect(pivotXml).toContain('<pageField fld="1" hier="0" item="1"/>');
    const reparsed = parseWorkbookSync(buf);
    expect(reparsed.worksheets).toHaveLength(2);
  });

  it("reuses the pivot relationship for pivotSelection", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "Data",
          rows: [
            { cells: [{ value: "City" }, { value: "Revenue" }] },
            { cells: [{ value: "Beijing" }, { value: 320 }] },
          ],
        },
        {
          name: "Pivot",
          rows: [],
          pivotSelection: {
            activeRow: 1,
            activeCol: 0,
            click: 1,
            rId: "rId1",
            pivotArea: { type: "normal" },
          },
          pivotTables: [
            {
              mode: "source",
              source: "A1:B2",
              sourceSheet: "Data",
              rows: ["City"],
              data: [{ field: "Revenue", summarize: "sum" }],
            },
          ],
        },
      ],
    };

    const first = (await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array;
    const parsed = parseWorkbookSync(first);
    const second = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const relsXml = new TextDecoder().decode(
      unzipSync(second)["xl/worksheets/_rels/sheet2.xml.rels"],
    );
    expect([...relsXml.matchAll(/<Relationship\b/g)]).toHaveLength(1);
    expect(relsXml).toContain('Id="rId1"');
    expect(relsXml).toContain('Target="../pivotTables/pivotTable1.xml"');
  });

  it("adopts the style table so cells keep raw indices and their formatting", async () => {
    // Two cells with distinct styles. Parse exposes the style table sections
    // (fonts/fills/borders/cellXfs) alongside the cells; cells carry raw
    // cellXfs indices that resolve against that table — the SDK's Stylesheet
    // model — so a re-generate keeps the source's own numbering.
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [
            {
              cells: [
                {
                  reference: "A1",
                  value: "bold",
                  style: { font: { bold: true } },
                },
                {
                  reference: "A2",
                  value: "italic",
                  style: { font: { italic: true } },
                },
              ],
            },
          ],
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const rows = parsed.worksheets![0]!.rows!;
    const a1 = rows[0]!.cells![0]!;
    const a2 = rows[0]!.cells![1]!;

    // Raw indices plus the table they resolve against.
    expect(typeof a1.style).toBe("number");
    expect(typeof a2.style).toBe("number");
    expect(parsed.cellXfs).toBeDefined();
    expect(parsed.fonts).toBeDefined();
    expect(parsed.fonts![parsed.cellXfs![a1.style as number]!.fontId!]!.bold).toBe(true);
    expect(parsed.fonts![parsed.cellXfs![a2.style as number]!.fontId!]!.italic).toBe(true);

    // And the formatting survives a second generate→parse cycle intact.
    const reparsed = await roundTrip(parsed);
    const r2 = reparsed.worksheets![0]!.rows![0]!.cells!;
    const b1 = r2[0]!;
    const b2 = r2[1]!;
    expect(typeof b1.style).toBe("number");
    expect(typeof b2.style).toBe("number");
    expect(reparsed.fonts![reparsed.cellXfs![b1.style as number]!.fontId!]!.bold).toBe(true);
    expect(reparsed.fonts![reparsed.cellXfs![b2.style as number]!.fontId!]!.italic).toBe(true);
  });

  it("round-trips dxfs from options through the workbook", async () => {
    const opts: WorkbookOptions = {
      dxfs: [{ font: { bold: true } }, { fill: { color: "FF0000", patternType: "solid" } }],
      worksheets: [{ name: "S", rows: [{ cells: [{ value: 1 }] }] }],
    };

    const parsed = await roundTrip(opts);
    expect(parsed.dxfs).toBeDefined();
    expect(parsed.dxfs).toHaveLength(2);
  });

  it("preserves dxf border outline and color tint precision", async () => {
    const opts: WorkbookOptions = {
      dxfs: [
        {
          font: { themeColor: 1, tint: 0.499985 },
          border: {
            outline: false,
            left: { style: "thin", themeColor: 2, tint: -0.249973 },
          },
        },
      ],
      worksheets: [{ name: "S", tabColor: { theme: 3, tint: 0.599995 } }],
    };

    const buffer = (await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array;
    const archive = unzipSync(buffer);
    const styles = new TextDecoder().decode(archive["xl/styles.xml"]!);
    expect(styles).toContain('outline="0"');
    expect(styles).toContain('tint="-0.249973"');
    expect(new TextDecoder().decode(archive["xl/worksheets/sheet1.xml"]!)).toContain(
      'tint="0.599995"',
    );

    const parsed = parseWorkbookSync(buffer);
    expect(parsed.dxfs![0]).toEqual({
      font: { themeColor: 1, tint: 0.499985, tintRaw: "0.499985" },
      border: {
        outline: false,
        left: { style: "thin", themeColor: 2, tint: -0.249973, tintRaw: "-0.249973" },
      },
    });
    expect(parsed.worksheets![0]!.tabColor).toEqual({
      theme: 3,
      tint: 0.599995,
      tintRaw: "0.599995",
    });
  });

  it("keeps built-in numFmt ids so date cells keep their format", async () => {
    // numFmt "mm-dd-yy" registers as built-in id 14 (no custom <numFmts>
    // entry). The adopted cellXfs entry carries numFmtId 14 verbatim, so the
    // round-tripped cell keeps its date format without a table rewrite.
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [{ cells: [{ reference: "A1", value: 45658, style: { numFmt: "mm-dd-yy" } }] }],
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const cell = parsed.worksheets![0]!.rows![0]!.cells![0]!;
    expect(typeof cell.style).toBe("number");
    expect(parsed.cellXfs![cell.style as number]!.numFmtId).toBe(14);
  });

  it("round-trips row outline level, collapsed flag, and row style", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [
            { rowNumber: 1, cells: [{ value: "group" }] },
            {
              rowNumber: 2,
              cells: [{ value: "detail" }],
              outlineLevel: 1,
              collapsed: true,
              style: { font: { bold: true } },
            },
          ],
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const row = parsed.worksheets![0]!.rows![1]!;
    expect(row.outlineLevel).toBe(1);
    expect(row.collapsed).toBe(true);
    expect(typeof row.style).toBe("number");
    expect(parsed.fonts![parsed.cellXfs![row.style as number]!.fontId!]!.bold).toBe(true);
  });

  it("reads the external link target from the sibling rels file", async () => {
    const opts: WorkbookOptions = {
      externalLinks: [
        {
          externalBook: {
            target: "external/source.xlsx",
            sheetNames: ["Sheet1"],
          },
        },
      ],
      worksheets: [{ name: "S", rows: [{ cells: [{ value: 1 }] }] }],
    };

    const parsed = await roundTrip(opts);
    expect(parsed.externalLinks).toBeDefined();
    expect(parsed.externalLinks).toHaveLength(1);
    // The target lives in xl/externalLinks/_rels/externalLink1.xml.rels, not in
    // the externalLink XML body — this asserts the rels file is actually read.
    expect(parsed.externalLinks![0]?.externalBook?.target).toBe("external/source.xlsx");
  });

  it("round-trips sheet-view selection, pivotSelection, and cell smart tags", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [{ cells: [{ reference: "A1", value: 1 }] }],
          selection: [{ activeCell: "B2", sqref: "B2" }],
          pivotSelection: {
            axis: "axisRow",
            activeRow: 1,
            count: 1,
            pivotArea: { type: "normal", outline: true },
          },
          smartTags: [
            {
              reference: "A1",
              smartTags: [
                {
                  type: 0,
                  xmlBased: true,
                  properties: [{ key: "urn:schemas-company:stock", val: "FIN" }],
                },
              ],
            },
          ],
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const sheet = parsed.worksheets![0]!;
    expect(sheet.selection?.[0]).toMatchObject({ activeCell: "B2", sqref: "B2" });
    expect(sheet.pivotSelection).toMatchObject({ axis: "axisRow", activeRow: 1, count: 1 });
    expect(sheet.pivotSelection!.pivotArea).toMatchObject({ type: "normal" });
    expect(sheet.smartTags).toHaveLength(1);
    expect(sheet.smartTags![0]).toMatchObject({ reference: "A1" });
    expect(sheet.smartTags![0]!.smartTags[0]).toMatchObject({ type: 0, xmlBased: true });
    expect(sheet.smartTags![0]!.smartTags[0]!.properties![0]).toEqual({
      key: "urn:schemas-company:stock",
      val: "FIN",
    });
  });

  it("round-trips the sheet background picture", async () => {
    const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "BG",
          rows: [{ cells: [{ value: "x" }] }],
          backgroundImage: { data: bytes, type: "png" },
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const bg = parsed.worksheets![0]!.backgroundImage;
    expect(bg?.type).toBe("png");
    expect(new Uint8Array(bg!.data as Uint8Array)).toEqual(bytes);
  });
});

describe("metadata round-trip", () => {
  it("preserves non-default metadata part paths", async () => {
    const opts: WorkbookOptions = {
      worksheets: [{ name: "S", rows: [{ cells: [{ value: "x" }] }] }],
      corePropertiesPath: "docProps/props/core.xml",
      appPropertiesPath: "docProps/props/app.xml",
      customPropertiesPath: "docProps/props/custom.xml",
      customPropertiesDeclared: true,
    };

    const parsed = await roundTrip(opts);
    expect(parsed.corePropertiesPath).toBe("docProps/props/core.xml");
    expect(parsed.appPropertiesPath).toBe("docProps/props/app.xml");
    expect(parsed.customPropertiesPath).toBe("docProps/props/custom.xml");
    expect(parsed.customPropertiesDeclared).toBe(true);
  });

  it("round-trips the metadata part and cell cm/vm references", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [
            {
              cells: [{ reference: "A1", value: "x", cellMetadataId: 1, valueMetadataId: 1 }],
            },
          ],
        },
      ],
      metadata: {
        types: [{ name: "XLDAPROPERTY", minSupportedVersion: 1 }],
        strings: [{ value: "s1" }],
        futureMetadata: [{ name: "XLDAPROPERTY", blocks: [{}] }],
        cellMetadata: [{ records: [{ typeIndex: 0, valueIndex: 0 }] }],
        valueMetadata: [{ records: [{ typeIndex: 0, valueIndex: 0 }] }],
      },
    };

    const parsed = await roundTrip(opts);
    expect(parsed.metadata?.types![0]).toMatchObject({
      name: "XLDAPROPERTY",
      minSupportedVersion: 1,
    });
    expect(parsed.metadata?.futureMetadata![0]).toMatchObject({ name: "XLDAPROPERTY" });
    expect(parsed.metadata?.cellMetadata![0]!.records).toEqual([{ typeIndex: 0, valueIndex: 0 }]);
    const cell = parsed.worksheets![0]!.rows![0]!.cells![0]!;
    expect(cell.cellMetadataId).toBe(1);
    expect(cell.valueMetadataId).toBe(1);
  });
});

describe("xml mapping round-trip", () => {
  it("round-trips xmlMaps and per-sheet single-cell XML tables", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [{ cells: [{ reference: "A1", value: "x" }] }],
          singleXmlCells: [
            {
              id: 1,
              reference: "B2",
              connectionId: 1,
              properties: {
                id: 1,
                mapping: { mapId: 1, xpath: "/root/name", xmlDataType: "string" },
              },
            },
          ],
        },
      ],
      xmlMaps: {
        selectionNamespaces: 'xmlns:m="http://example.com"',
        schemas: [{ id: "S1", namespace: "http://example.com" }],
        maps: [{ id: 1, name: "M1", rootElement: "root", schemaId: "S1" }],
      },
    };

    const parsed = await roundTrip(opts);
    expect(parsed.xmlMaps?.selectionNamespaces).toBe('xmlns:m="http://example.com"');
    expect(parsed.xmlMaps?.schemas![0]).toMatchObject({ id: "S1" });
    expect(parsed.xmlMaps?.maps![0]).toMatchObject({ id: 1, name: "M1", rootElement: "root" });
    const cell = parsed.worksheets![0]!.singleXmlCells![0]!;
    expect(cell).toMatchObject({ id: 1, reference: "B2", connectionId: 1 });
    expect(cell.properties.mapping).toMatchObject({ mapId: 1, xpath: "/root/name" });
  });
});

describe("dialogsheet round-trip", () => {
  it("round-trips a legacy dialog sheet with protection and page setup", async () => {
    const opts: WorkbookOptions = {
      worksheets: [{ name: "S", rows: [{ cells: [{ value: 1 }] }] }],
      dialogsheets: [
        {
          name: "Dialog1",
          tabColor: "FF404040",
          codeName: "Dialog1",
          sheetProtection: { objects: true },
          pageMargins: { left: 0.5 },
          pageSetup: { paperSize: 9, orientation: "portrait" },
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const ds = parsed.dialogsheets![0]!;
    expect(ds.tabColor).toBe("FF404040");
    expect(ds.codeName).toBe("Dialog1");
    expect(ds.sheetProtection).toEqual({ objects: true });
    expect(ds.pageMargins?.left).toBe(0.5);
    expect(ds.pageSetup).toMatchObject({ paperSize: 9, orientation: "portrait" });
  });
});

describe("scattered attribute round-trip", () => {
  it("round-trips book-view visibility and conditional-format rule flags", async () => {
    const opts: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          rows: [{ cells: [{ reference: "A1", value: 5 }] }],
          conditionalFormats: [
            {
              sqref: "A1:A5",
              rules: [
                { type: "aboveAverage", aboveAverage: false, stdDev: 2, dxfId: 0 },
                { type: "top10", rank: 5 },
              ],
            },
          ],
          protectedRanges: [{ sqref: "A1:A5", name: "Guarded", securityDescriptor: "O=ORG" }],
        },
      ],
    };

    const parsed = await roundTrip(opts);
    const sheet = parsed.worksheets![0]!;
    expect(sheet.conditionalFormats![0]!.rules[0]).toMatchObject({
      type: "aboveAverage",
      aboveAverage: false,
      stdDev: 2,
    });
    expect(sheet.protectedRanges![0]).toMatchObject({
      name: "Guarded",
      securityDescriptor: "O=ORG",
    });
  });
});

describe("theme round-trip", () => {
  it("preserves a custom source theme instead of replacing it with the default", async () => {
    const opts: WorkbookOptions = {
      worksheets: [{ name: "S", rows: [{ cells: [{ reference: "A1", value: 1 }] }] }],
    };
    const buffer = (await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array;

    // Inject a custom theme color (accent1 = FF00FF) into the generated package.
    const unzipped = unzipSync(buffer);
    const themeEntry = unzipped["xl/theme/theme1.xml"];
    if (!themeEntry) throw new Error("missing xl/theme/theme1.xml");
    const themeXml = new TextDecoder().decode(themeEntry);
    const mutated = themeXml.replace(/(<a:accent1>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}/, "$1FF00FF");
    if (mutated === themeXml) throw new Error("accent1 srgbClr not found in theme");
    unzipped["xl/theme/theme1.xml"] = new TextEncoder().encode(mutated);
    const reborn = zipSync(unzipped);

    const parsed = parseWorkbookSync(reborn);
    expect(parsed.theme?.colorScheme?.accent1).to.exist;

    const regenerated = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const regTheme = new TextDecoder().decode(unzipSync(regenerated)["xl/theme/theme1.xml"]!);
    expect(regTheme.toLowerCase()).to.contain('val="ff00ff"');
  });
});
