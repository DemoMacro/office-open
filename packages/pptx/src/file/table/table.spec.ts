import { describe, expect, it } from "vite-plus/test";

import { DrawingTable } from "./table";

const context = { stack: [] };

// Helper: minimal row with N text cells
function makeRows(...cellTexts: string[][]) {
  return {
    rows: cellTexts.map((texts) => ({
      cells: texts.map((t) => ({ text: t })),
    })),
  };
}

describe("DrawingTable", () => {
  describe("toXml", () => {
    it("basic table produces valid XML", () => {
      const xml = new DrawingTable(makeRows(["A1", "B1"])).toXml(context);
      expect(xml).toContain("<a:tbl>");
      expect(xml).toContain("</a:tbl>");
      expect(xml).toContain("<a:tblPr/>");
      expect(xml).toContain("<a:tblGrid>");
      expect(xml).toContain("<a:tr");
    });

    it("with columnWidths produces gridCol with width", () => {
      const xml = new DrawingTable({
        ...makeRows(["A", "B"]),
        columnWidths: [100, 200],
      }).toXml(context);
      expect(xml).toContain('<a:gridCol w="100"/>');
      expect(xml).toContain('<a:gridCol w="200"/>');
    });

    it("with bandRow produces tblPr attribute", () => {
      const xml = new DrawingTable({
        ...makeRows(["A"]),
        bandRow: true,
      }).toXml(context);
      expect(xml).toContain('bandRow="1"');
    });

    it("cell text content appears in XML", () => {
      const xml = new DrawingTable(makeRows(["Hello"])).toXml(context);
      expect(xml).toContain("<a:t>Hello</a:t>");
    });
  });
});
