import { describe, expect, it } from "vite-plus/test";

import { DocxWriteContext } from "./context";

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
