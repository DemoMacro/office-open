import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { DocxWriteContext } from "./context";
import { generateDocumentSync } from "./generate";

const CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_EX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

describe("DocxWriteContext passthrough relationships", () => {
  it("keeps opaque relationships under compiler-owned directories", () => {
    const ctx = new DocxWriteContext({
      sections: [],
      passthroughRelationships: [
        {
          source: "word/document.xml",
          relationshipType: CHART_EX_REL,
          target: "charts/chartEx1.xml",
          rId: "rId8",
        },
        {
          source: "word/document.xml",
          relationshipType: CHART_REL,
          target: "charts/chart1.xml",
          rId: "rId9",
        },
      ],
    });

    ctx.addPassthroughDocumentRelationships();
    const xml = ctx.document.relationships.serialize();
    expect(xml).toContain(`Type="${CHART_EX_REL}" Target="charts/chartEx1.xml"`);
    expect(xml).toContain(`Type="${CHART_REL}" Target="charts/chart1.xml"`);
  });

  it("deduplicates only the same relationship kind and target", () => {
    const ctx = new DocxWriteContext({
      sections: [],
      passthroughRelationships: [
        {
          source: "word/document.xml",
          relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/theme",
          target: "theme/theme1.xml",
          rId: "rId8",
        },
        {
          source: "word/document.xml",
          relationshipType: CHART_REL,
          target: "charts/chart1.xml",
          rId: "rId9",
        },
        {
          source: "word/document.xml",
          relationshipType: CHART_REL,
          target: "charts/chart2.xml",
          rId: "rId10",
        },
      ],
    });

    ctx.document.relationships.addRelationship(20, CHART_REL, "charts/chart1.xml");
    ctx.addPassthroughDocumentRelationships();

    const xml = ctx.document.relationships.serialize();
    expect(xml.match(/Target="theme\/theme1\.xml"/g)).toHaveLength(1);
    expect(xml.match(/Target="charts\/chart1\.xml"/g)).toHaveLength(1);
    expect(xml.match(/Target="charts\/chart2\.xml"/g)).toHaveLength(1);
  });
});

describe("round-tripped styles without latentStyles", () => {
  it("does not inject the factory latent-style table", () => {
    // Word 2010 transitional documents can omit w:latentStyles entirely; the
    // factory fallback would add 20 lsdException entries the source lacks.
    const ctx = new DocxWriteContext({
      sections: [],
      styles: {
        roundTripped: true,
        paragraphStyles: [{ id: "Normal", name: "Normal", default: true }],
      },
    });
    const xml = ctx.styles.serialize();
    expect(xml).not.toContain("w:latentStyles");
    expect(xml).not.toContain("w:lsdException");
  });
});

describe("generation with passthrough source ids and a source-new part", () => {
  it("keeps source ids for re-used parts and allocates the new part above them", () => {
    // Word-typical layout: theme=rId6, fontTable=rId7. The edited document
    // adds comments — a part the source didn't carry. Its auto-allocated id
    // must land above the source id space, or fontTable's source re-use
    // collides with theme and the package is corrupted for Office.
    const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    const FONT_TABLE_REL =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    const passthroughRelationships = [
      {
        source: "word/document.xml",
        relationshipType: THEME_REL,
        target: "theme/theme1.xml",
        rId: "rId6",
      },
      {
        source: "word/document.xml",
        relationshipType: FONT_TABLE_REL,
        target: "fontTable.xml",
        rId: "rId7",
      },
    ];
    const rawParts = [
      {
        path: "word/theme/theme1.xml",
        data: new TextEncoder().encode("<?xml version='1.0'?><theme/>"),
      },
    ];
    const buffer = generateDocumentSync({
      sections: [{ children: [{ paragraph: { text: "probe" } }] }],
      comments: [
        {
          id: 1,
          author: "P",
          initials: "P",
          date: "2026-09-01T00:00:00Z",
          children: [{ text: "probe comment" }],
        },
      ],
      passthroughRelationships,
      rawParts,
    }) as Uint8Array;
    const rels = new TextDecoder().decode(unzipSync(buffer)["word/_rels/document.xml.rels"]);
    const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(rels).toMatch(/Id="rId6"[^>]*theme\/theme1\.xml/);
    expect(rels).toMatch(/Id="rId7"[^>]*fontTable\.xml/);
    // Comments — absent from the source — must land above the source id space
    const commentsId = Number(/Id="rId(\d+)"[^>]*relationships\/comments"/.exec(rels)?.[1]);
    expect(commentsId).toBeGreaterThan(7);
  });
});
