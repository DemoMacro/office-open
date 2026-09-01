import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./generate";
import type { WorkbookOptions } from "./parts/file";

// A rebuilt worksheet rels table mixes model allocations (tables, comments, …)
// with passthrough re-emission at source ids (verbatim references in the sheet
// XML — drawing, printerSettings — keep working only if the ids stay). The
// source id space is reserved up front and every model allocation flows
// through the rels' own watermark, so a preserved source id can never be
// handed out twice.

const PRINTER_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

const decodeEntry = (buffer: Uint8Array, path: string): string => {
  const unzipped = unzipSync(buffer);
  const entry = unzipped[path];
  if (!entry) throw new Error(`missing zip entry: ${path}`);
  return new TextDecoder().decode(entry);
};

const table = (ref: string) => ({
  displayName: `T${ref}`,
  name: `T${ref}`,
  ref,
  columns: [{ name: "A" }, { name: "B" }],
});

// Source sheet1.xml.rels: rId2 a verbatim drawing part, rId4 printerSettings
// (pageSetup references it by r:id). Four tables rebuild the rels — without
// the reserve their allocations collide with the preserved source ids.
describe("worksheet rels with passthrough source ids and rebuilt parts", () => {
  it("keeps source ids for re-used rels and allocates tables above them", async () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          drawingRid: "rId2",
          pageSetup: { printerSettingsRId: "rId4" },
          tables: [table("A1:B4"), table("D1:E4"), table("G1:H4"), table("J1:K4")],
        },
      ],
      passthroughRelationships: [
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: DRAWING_REL,
          target: "../drawings/drawing1.xml",
          rId: "rId2",
        },
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: PRINTER_REL,
          target: "../printerSettings/printerSettings1.bin",
          rId: "rId4",
        },
      ],
      rawParts: [
        { path: "xl/drawings/drawing1.xml", data: "<xdr:wsDr/>" },
        { path: "xl/printerSettings/printerSettings1.bin", data: new Uint8Array([1, 2, 3]) },
      ],
    };

    const buffer = await generateWorkbook(options);
    const rels = decodeEntry(buffer, "xl/worksheets/_rels/sheet1.xml.rels");

    const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    // Both verbatim references keep their source ids
    expect(rels).toMatch(new RegExp(`Id="rId2"[^>]*Type="${DRAWING_REL}"`));
    expect(rels).toMatch(new RegExp(`Id="rId4"[^>]*Type="${PRINTER_REL}"`));
    // The four fresh tables land above the source id space
    const tableIds = [...rels.matchAll(/Id="rId(\d+)"[^>]*relationships\/table"/g)].map((m) =>
      Number(m[1]),
    );
    expect(tableIds).toHaveLength(4);
    expect(tableIds.every((id) => id > 4)).toBe(true);
  });
});
