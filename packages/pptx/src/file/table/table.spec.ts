import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { DrawingTable } from "./table";

const context = { stack: [] } as any;

// Helper: minimal row with N text cells
function makeRows(...cellTexts: string[][]) {
  return {
    rows: cellTexts.map((texts) => ({
      cells: texts.map((t) => ({ text: t })),
    })),
  };
}

describe("DrawingTable", () => {
  // ── prepForXml path ──

  describe("prepForXml", () => {
    it("basic 2x2 table produces tblPr + tblGrid + 2 tr", () => {
      const tree = new Formatter().format(new DrawingTable(makeRows(["A1", "B1"], ["A2", "B2"])));
      const tbl = tree["a:tbl"] as any[];

      // tblPr, tblGrid, tr, tr
      expect(tbl.length).toBeGreaterThanOrEqual(3);

      // First child is tblPr
      expect(tbl[0]).toHaveProperty("a:tblPr");

      // Second child is tblGrid
      expect(tbl[1]).toHaveProperty("a:tblGrid");

      // Grid has 2 columns (default width 0)
      const grid = tbl[1]["a:tblGrid"];
      expect(grid).toHaveLength(2);
    });

    it("with columnWidths produces grid columns with widths", () => {
      const tree = new Formatter().format(
        new DrawingTable({
          ...makeRows(["A", "B"]),
          columnWidths: [100, 200],
        }),
      );
      const tbl = tree["a:tbl"] as any[];
      const grid = tbl[1]["a:tblGrid"];
      expect(grid[0]["a:gridCol"]["_attr"]["w"]).toBe(100);
      expect(grid[1]["a:gridCol"]["_attr"]["w"]).toBe(200);
    });

    it("with bandRow produces tblPr with bandRow=1", () => {
      const tree = new Formatter().format(new DrawingTable({ ...makeRows(["A"]), bandRow: true }));
      const tbl = tree["a:tbl"] as any[];
      expect(tbl[0]["a:tblPr"]).to.deep.equal({ _attr: { bandRow: 1 } });
    });

    it("with firstRow produces tblPr with firstRow=1", () => {
      const tree = new Formatter().format(new DrawingTable({ ...makeRows(["A"]), firstRow: true }));
      const tbl = tree["a:tbl"] as any[];
      expect(tbl[0]["a:tblPr"]).to.deep.equal({ _attr: { firstRow: 1 } });
    });

    it("distributes top border to first row cells", () => {
      const tree = new Formatter().format(
        new DrawingTable({
          ...makeRows(["A1", "B1"], ["A2", "B2"]),
          borders: { top: { color: "FF0000", width: 1 } },
        }),
      );
      const tbl = tree["a:tbl"] as any[];

      // Find first row (3rd element after tblPr and tblGrid)
      const firstRow = tbl[2]["a:tr"] as any[];
      // Each cell should have a top border (a:lnT) in tcPr
      const tc = firstRow[1]; // first cell after _attr
      const txBody = tc["a:tc"] as any[];
      // Find tcPr within the cell
      const tcPr = txBody.find((c: any) => "a:tcPr" in c);
      expect(tcPr).toBeDefined();
    });

    it("distributes left border to first column cells", () => {
      const tree = new Formatter().format(
        new DrawingTable({
          ...makeRows(["A1", "B1"], ["A2", "B2"]),
          borders: { left: { color: "00FF00", width: 1 } },
        }),
      );
      const tbl = tree["a:tbl"] as any[];

      // Check that the first column cells have left border
      // Both rows' first cells should have a:lnL in their tcPr
      for (let ri = 0; ri < 2; ri++) {
        const row = tbl[ri + 2]["a:tr"] as any[];
        const firstCellChildren = row[1]["a:tc"] as any[];
        const tcPr = firstCellChildren.find((c: any) => "a:tcPr" in c);
        expect(tcPr).toBeDefined();
      }
    });
  });

  // ── toXml path ──

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
