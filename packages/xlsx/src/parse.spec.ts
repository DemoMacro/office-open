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
    const rows = parsed.worksheets![0].rows!;
    const a1 = rows[0].cells![0];
    const a2 = rows[0].cells![1];

    // Resolved style objects must be present (not just raw indices).
    expect(a1.style).toBeDefined();
    expect(a2.style).toBeDefined();
    expect(a1.style!.font?.bold).toBe(true);
    expect(a2.style!.font?.italic).toBe(true);

    // And the formatting survives a second generate→parse cycle intact.
    const reparsed = await roundTrip(parsed);
    const r2 = reparsed.worksheets![0].rows![0].cells!;
    expect(r2[0].style!.font?.bold).toBe(true);
    expect(r2[1].style!.font?.italic).toBe(true);
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
    expect(parsed.externalLinks![0].externalBook?.target).toBe("external/source.xlsx");
  });
});
