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
const VML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const OLE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const CONTROLS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/controls";
const CUSTOMXML_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

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

  // Source ids the sheet XML emits verbatim beyond drawing/printerSettings —
  // OLE objects, controls, header/footer VML, custom parts — resolve before
  // stringify, so when the model's allocations took the source id the XML
  // reference follows the renumbered rel instead of dangling.
  it("remaps verbatim rid fields when the model took their source ids", async () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "S",
          drawingRid: "rId2",
          legacyDrawingHF: "rId5",
          oleObjects: [{ shapeId: 1, rId: "rId6", properties: { iconRid: "rId7" } }],
          controls: [{ shapeId: 2, rId: "rId8", iconRid: "rId9" }],
          customProperties: [{ name: "prop", rId: "rId10" }],
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
          relationshipType: VML_REL,
          target: "../drawings/vmlDrawing1.vml",
          rId: "rId5",
        },
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: OLE_REL,
          target: "../embeddings/oleObject1.bin",
          rId: "rId6",
        },
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: IMAGE_REL,
          target: "../media/image1.png",
          rId: "rId7",
        },
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: CONTROLS_REL,
          target: "../ctrlProps/ctrlProp1.xml",
          rId: "rId8",
        },
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: IMAGE_REL,
          target: "../media/image2.png",
          rId: "rId9",
        },
        {
          source: "xl/worksheets/sheet1.xml",
          relationshipType: CUSTOMXML_REL,
          target: "../customXml/item1.xml",
          rId: "rId10",
        },
      ],
      rawParts: [
        { path: "xl/drawings/drawing1.xml", data: "<xdr:wsDr/>" },
        { path: "xl/drawings/vmlDrawing1.vml", data: "<xml/>" },
        { path: "xl/embeddings/oleObject1.bin", data: new Uint8Array([1]) },
        { path: "xl/media/image1.png", data: new Uint8Array([2]) },
        { path: "xl/ctrlProps/ctrlProp1.xml", data: "<ctrlProp/>" },
        { path: "xl/media/image2.png", data: new Uint8Array([3]) },
        { path: "xl/customXml/item1.xml", data: "<item/>" },
      ],
    };

    const buffer = await generateWorkbook(options);
    const rels = decodeEntry(buffer, "xl/worksheets/_rels/sheet1.xml.rels");
    const sheet = decodeEntry(buffer, "xl/worksheets/sheet1.xml");

    const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    // Free source ids keep their slot; every verbatim reference in the XML
    // resolves to a declared rel id.
    expect(rels).toContain(`Id="rId5"`);
    for (const rid of ["rId5", "rId6", "rId7", "rId8", "rId9", "rId10"]) {
      const declared = rels.includes(`Id="${rid}"`);
      const referenced = sheet.includes(`"${rid}"`);
      if (referenced) expect(declared).toBe(true);
    }
    // The verbatim fields all appear with resolvable ids
    expect(sheet).toMatch(/<legacyDrawingHF r:id="rId5"\/>/);
    expect(sheet).toMatch(/<oleObject [^>]*r:id="rId6"/);
    expect(sheet).toMatch(/<objectPr[^>]*r:id="rId7"/);
    expect(sheet).toMatch(/<control [^>]*r:id="rId8"/);
    expect(sheet).toMatch(/<controlPr[^>]*r:id="rId9"/);
    expect(sheet).toMatch(/<customPr name="prop" r:id="rId10"\/>/);
  });
});
