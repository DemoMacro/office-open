import { unzipSync, zipSync } from "@office-open/core";
import type { WorkbookOptions } from "@parts/file";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./generate";
import { parseWorkbook } from "./parse";

// Full-file round-trip: generateWorkbook → Buffer → parseWorkbook → WorkbookOptions.
// Proves the three previously-deferred parse gaps are resolved on the read path.

async function roundTrip(opts: WorkbookOptions): Promise<WorkbookOptions> {
  const buf = (await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array;
  return parseWorkbook(buf);
}

describe("parseWorkbook round-trip", () => {
  it("resolves cell.style so a fresh Styles table keeps the right formatting", async () => {
    // Two cells with distinct styles. After parse→regenerate the Styles table
    // is rebuilt from scratch (indices may differ), so cell.styleIndex alone
    // would point at the wrong xf. The fix resolves cell.style instead.
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
    const a1 = rows[0]!.cells![0];
    const a2 = rows[0]!.cells![1];

    // Resolved style objects must be present (not just raw indices).
    expect(a1?.style).toBeDefined();
    expect(a2?.style).toBeDefined();
    expect(a1!.style!.font?.bold).toBe(true);
    expect(a2!.style!.font?.italic).toBe(true);

    // And the formatting survives a second generate→parse cycle intact.
    const reparsed = await roundTrip(parsed);
    const r2 = reparsed.worksheets![0]!.rows![0]!.cells!;
    expect(r2[0]!.style!.font?.bold).toBe(true);
    expect(r2[1]!.style!.font?.italic).toBe(true);
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
          selection: { activeCell: "B2", sqref: "B2" },
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
    expect(sheet.selection).toMatchObject({ activeCell: "B2", sqref: "B2" });
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
});

describe("metadata round-trip", () => {
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
        cellMetadata: [{ records: [{ t: 0, v: 0 }] }],
        valueMetadata: [{ records: [{ t: 0, v: 0 }] }],
      },
    };

    const parsed = await roundTrip(opts);
    expect(parsed.metadata?.types![0]).toMatchObject({
      name: "XLDAPROPERTY",
      minSupportedVersion: 1,
    });
    expect(parsed.metadata?.futureMetadata![0]).toMatchObject({ name: "XLDAPROPERTY" });
    expect(parsed.metadata?.cellMetadata![0]!.records).toEqual([{ t: 0, v: 0 }]);
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
              r: "B2",
              connectionId: 1,
              xmlCellPr: { id: 1, xmlPr: { mapId: 1, xpath: "/root/name", xmlDataType: "string" } },
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
    expect(cell).toMatchObject({ id: 1, r: "B2", connectionId: 1 });
    expect(cell.xmlCellPr.xmlPr).toMatchObject({ mapId: 1, xpath: "/root/name" });
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

    const parsed = parseWorkbook(reborn);
    expect(parsed.theme?.colorScheme?.accent1).to.exist;

    const regenerated = (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array;
    const regTheme = new TextDecoder().decode(unzipSync(regenerated)["xl/theme/theme1.xml"]!);
    expect(regTheme.toLowerCase()).to.contain('val="ff00ff"');
  });
});
