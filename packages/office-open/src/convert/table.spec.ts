import type { TableOptions as DocxTableOptions } from "@office-open/docx";
import type { TableOptions as PptxTableOptions } from "@office-open/pptx";
import { describe, expect, it } from "vitest";

import { toDocxTable, toPptxTable, toXlsxTable } from "./table";

const docxTable: DocxTableOptions = {
  rows: [
    {
      cells: [
        { children: [{ paragraph: { children: [{ text: "A1" }] } }] },
        { children: [{ paragraph: { children: [{ text: "B1" }] } }] },
      ],
    },
  ],
  columnWidths: [1000, 2000], // twip
  firstRow: true,
  bandRow: true,
};

const pptxTable: PptxTableOptions = {
  rows: [{ cells: [{ text: "A1" }, { text: "B1" }] }],
  columnWidths: [1000, 2000], // EMU
  firstRow: true,
};

describe("toPptxTable (docx → pptx)", () => {
  it("passes base flags through and converts column widths twip→EMU", () => {
    const pptx = toPptxTable(docxTable);
    expect(pptx.firstRow).toBe(true);
    expect(pptx.bandRow).toBe(true);
    expect(pptx.columnWidths).toEqual([635000, 1270000]); // ×635
  });

  it("bridges cell content w:p → a:p", () => {
    const cell = toPptxTable(docxTable).rows[0]!.cells[0]!;
    expect(cell.children).toHaveLength(1);
  });
});

describe("toDocxTable (pptx → docx)", () => {
  it("passes base flags through and converts column widths EMU→twip", () => {
    const docx = toDocxTable(pptxTable);
    expect(docx.firstRow).toBe(true);
    expect(docx.columnWidths).toEqual([2, 3]); // round(1000/635), round(2000/635)
  });

  it("bridges cell text → w:p", () => {
    const row = toDocxTable(pptxTable).rows[0]!;
    if (!("cells" in row)) throw new Error("expected plain row");
    const cell = row.cells[0]!;
    if (!("children" in cell)) throw new Error("expected plain cell");
    expect(cell.children).toHaveLength(1);
  });

  it("maps pptx solid fill → docx shading", () => {
    const pptx: PptxTableOptions = {
      rows: [{ cells: [{ text: "x", fill: { type: "solid", color: "FF0000" } }] }],
    };
    const row = toDocxTable(pptx).rows[0]!;
    if (!("cells" in row)) throw new Error("expected plain row");
    const cell = row.cells[0]!;
    if (!("children" in cell)) throw new Error("expected plain cell");
    expect(cell.shading).toEqual({ fill: "FF0000" });
  });

  it("maps pptx scheme color → docx themeColor", () => {
    const pptx: PptxTableOptions = {
      rows: [{ cells: [{ text: "x", fill: { type: "solid", color: { value: "accent1" } } }] }],
    };
    const row = toDocxTable(pptx).rows[0]!;
    if (!("cells" in row)) throw new Error("expected plain row");
    const cell = row.cells[0]!;
    if (!("children" in cell)) throw new Error("expected plain cell");
    expect(cell.shading).toEqual({ themeColor: "accent1" });
  });
});

describe("toXlsxTable (docx → xlsx visual restoration)", () => {
  it("extracts cell value, merges from span, and column widths", () => {
    const docx: DocxTableOptions = {
      rows: [
        {
          cells: [
            { children: [{ paragraph: { children: [{ text: "merged" }] } }], columnSpan: 2 },
            { children: [{ paragraph: { children: [{ text: "C1" }] } }] },
          ],
        },
      ],
      columnWidths: [1000, 2000],
    };
    const xlsx = toXlsxTable(docx);
    expect(xlsx.rows[0]!.cells![0]!.value).toBe("merged");
    expect(xlsx.mergeCells).toEqual([{ ref: "A1:B1" }]);
    expect(xlsx.columns).toHaveLength(2);
    expect(xlsx.columns![0]!.width!).toBeCloseTo(8.78, 1); // 635000 EMU / 72325
  });
});

describe("xlsx visual table source coordinates", () => {
  it("omits covered cells behind xlsx merge anchors", () => {
    const source = {
      mergeCells: [{ ref: "A1:B2" }],
      rows: [
        {
          rowNumber: 1,
          cells: [
            { reference: "A1", value: "merged" },
            { reference: "B1", value: "covered" },
          ],
        },
        {
          rowNumber: 2,
          cells: [
            { reference: "A2", value: "covered" },
            { reference: "B2", value: "covered" },
          ],
        },
      ],
    };
    const docx = toDocxTable(source);
    const docxAnchor = docx.rows[0]!;
    if (!("cells" in docxAnchor)) throw new Error("expected plain row");
    expect(docxAnchor.cells).toHaveLength(1);
    expect(docxAnchor.cells[0]).toMatchObject({ columnSpan: 2, rowSpan: 2 });
    expect(docx.rows[1]!).toEqual({ cells: [] });

    const pptx = toPptxTable(source);
    expect(pptx.rows[0]!.cells).toHaveLength(1);
    expect(pptx.rows[0]!.cells[0]).toMatchObject({ columnSpan: 2, rowSpan: 2 });
    expect(pptx.rows[1]!.cells).toEqual([]);
  });

  it("preserves sparse physical rows and merge spans", () => {
    const docx = toDocxTable({
      mergeCells: [{ ref: "A1:B2" }],
      rows: [
        { rowNumber: 1, cells: [{ reference: "A1", value: "merged" }] },
        { rowNumber: 3, cells: [{ reference: "C3", value: "C3" }] },
      ],
    });
    expect(docx.rows).toHaveLength(3);
    const anchor = docx.rows[0]!;
    if (!("cells" in anchor)) throw new Error("expected plain row");
    expect(anchor.cells[0]).toMatchObject({ columnSpan: 2, rowSpan: 2 });
    const sparse = docx.rows[2]!;
    if (!("cells" in sparse)) throw new Error("expected plain row");
    const cell = sparse.cells[2]!;
    if (!("children" in cell)) throw new Error("expected plain cell");
    expect(cell.children[0]).toEqual({
      paragraph: { children: [{ text: "C3" }] },
    });
  });
});

describe("round-trip docx → pptx → docx", () => {
  it("preserves base flags and column widths (635 divides evenly)", () => {
    const back = toDocxTable(toPptxTable(docxTable));
    expect(back.firstRow).toBe(true);
    expect(back.bandRow).toBe(true);
    expect(back.columnWidths).toEqual([1000, 2000]);
  });
});
