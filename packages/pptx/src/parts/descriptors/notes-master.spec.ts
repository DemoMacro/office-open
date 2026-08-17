import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { PptxWriteContext } from "../../context";
import { notesMasterDesc } from "./notes-master";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as PptxWriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(xml: string) {
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  const opts = notesMasterDesc.parse(el, readCtx);
  return { opts, xml: notesMasterDesc.stringify(opts, writeCtx)! };
}

describe("notesMasterDesc", () => {
  it("emits the default fresh layout (bgRef 1001 + 9-level notes style)", () => {
    const xml = notesMasterDesc.stringify({}, writeCtx)!;
    expect(xml).toContain('<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>');
    expect(xml).toContain('<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0"');
    expect(xml).toContain('<a:defRPr sz="1200" kern="1200">');
    expect(xml).toContain('<a:lvl9pPr marL="3657600"');
  });

  it("round-trips background, hf, colorMapping, and notesStyle levels", () => {
    const source =
      '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:bg><p:bgRef idx="1002"><a:srgbClr val="4472C4"/></p:bgRef></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      "<p:grpSpPr/></p:spTree></p:cSld>" +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:hf dt="1" hdr="0" ftr="1" sldNum="0"/>' +
      '<p:notesStyle><a:lvl1pPr marL="0"><a:defRPr sz="1400"/></a:lvl1pPr>' +
      '<a:lvl2pPr marL="457200"><a:defRPr sz="1200"/></a:lvl2pPr></p:notesStyle>' +
      "</p:notesMaster>";

    const { opts, xml } = roundTrip(source);
    expect(opts.background?.reference).toEqual({ index: 1002, color: { value: "4472C4" } });
    expect(opts.headerFooter).toEqual({
      date: true,
      header: false,
      footer: true,
      slideNumber: false,
    });
    expect(opts.notesStyle?.levels?.[0]?.defaultRun?.size).toBe(14);
    expect(opts.notesStyle?.levels?.[1]?.defaultRun?.size).toBe(12);
    // Regeneration keeps the parsed values (idempotent round-trip).
    expect(xml).toContain('<p:bgRef idx="1002"><a:srgbClr val="4472C4"/></p:bgRef>');
    expect(xml).toContain('<p:hf dt="1" hdr="0" ftr="1" sldNum="0"/>');
    expect(xml).toContain('<a:lvl1pPr marL="0"><a:defRPr sz="1400">');
  });

  it("emits custom spTree children with explicit ids", () => {
    const xml = notesMasterDesc.stringify(
      {
        children: [
          {
            shape: { id: 5, name: "Note badge", x: 0, y: 0, width: 100, height: 100 },
          },
        ],
      },
      writeCtx,
    )!;
    expect(xml).toContain('<p:cNvPr id="5" name="Note badge"');
  });
});
