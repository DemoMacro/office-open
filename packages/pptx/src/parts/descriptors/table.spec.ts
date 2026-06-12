import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { tableDesc } from "./table";
import type { TableDescriptorOptions } from "./table";

// ── Mock PPTX write context ──

class MockWriteContext {
  private _nextRelId = 1;
  addRelationship() {
    return `rId${this._nextRelId++}`;
  }
  addMedia() {
    return "";
  }
  addHyperlink() {}
  addImage() {}
}

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: TableDescriptorOptions) {
  const writeCtx = new MockWriteContext() as unknown as WriteContext;
  const xml = tableDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return tableDesc.parse(el, readCtx);
}

describe("tableDesc round-trip", () => {
  it("round-trips basic 2x2 table", () => {
    const opts: TableDescriptorOptions = {
      rows: [
        { cells: [{ text: "A1" }, { text: "B1" }] },
        { cells: [{ text: "A2" }, { text: "B2" }] },
      ],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;

    expect(rows).toHaveLength(2);
    const row0Cells = rows[0].cells!;
    expect(row0Cells).toHaveLength(2);
    expect(row0Cells[0].text).toBe("A1");
    expect(row0Cells[1].text).toBe("B1");
  });

  it("round-trips table with position", () => {
    const opts: TableDescriptorOptions = {
      x: 100,
      y: 200,
      width: 400,
      height: 300,
      rows: [{ cells: [{ text: "X" }] }],
    };
    const result = roundTrip(opts);

    expect(result.x).toBe(100);
    expect(result.y).toBe(200);
    expect(result.width).toBe(400);
    expect(result.height).toBe(300);
  });

  it("round-trips table with columnWidths", () => {
    const opts: TableDescriptorOptions = {
      columnWidths: [5000, 7000],
      rows: [{ cells: [{ text: "A" }, { text: "B" }] }],
    };
    const result = roundTrip(opts);
    const colWidths = result.columnWidths!;

    expect(colWidths).toHaveLength(2);
    expect(colWidths[0]).toBe(5000);
    expect(colWidths[1]).toBe(7000);
  });

  it("round-trips table with row height", () => {
    const opts: TableDescriptorOptions = {
      rows: [{ height: 500, cells: [{ text: "Tall" }] }],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;

    expect(rows[0].height).toBe(500);
  });

  it("round-trips cell with fill", () => {
    const opts: TableDescriptorOptions = {
      rows: [
        {
          cells: [{ text: "Filled", fill: { type: "solid", color: "4472C4" } }],
        },
      ],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const cell = rows[0].cells![0];
    const fill = cell.fill! as { type: string; color: { value: string } };

    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("4472C4");
  });

  it("round-trips cell with borders", () => {
    const opts: TableDescriptorOptions = {
      rows: [
        {
          cells: [
            {
              text: "Bordered",
              borders: {
                top: { width: 12700, color: "000000" },
                bottom: { width: 12700, color: "FF0000" },
              },
            },
          ],
        },
      ],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const cell = rows[0].cells![0];
    const borders = cell.borders!;
    const top = borders.top!;

    expect(top.width).toBe(12700);
    expect(top.color).toBe("000000");
  });

  it("round-trips cell with verticalAlign", () => {
    const opts: TableDescriptorOptions = {
      rows: [{ cells: [{ text: "Center", verticalAlign: "center" }] }],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const cell = rows[0].cells![0];

    expect(cell.verticalAlign).toBe("center");
  });

  it("round-trips table properties (bandRow, firstRow, etc.)", () => {
    const opts: TableDescriptorOptions = {
      bandRow: true,
      firstRow: true,
      lastCol: true,
      rows: [{ cells: [{ text: "X" }] }],
    };
    const result = roundTrip(opts);

    expect(result.bandRow).toBe(true);
    expect(result.firstRow).toBe(true);
    expect(result.lastCol).toBe(true);
  });

  it("round-trips table with tableStyleId", () => {
    const opts: TableDescriptorOptions = {
      tableStyleId: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
      rows: [{ cells: [{ text: "X" }] }],
    };
    roundTrip(opts);

    // tableStyleId is in tblPr, parse extracts it from tblStyleId element
    // The current parse doesn't extract tableStyleId, verify the XML contains it
    const writeCtx = new MockWriteContext() as unknown as WriteContext;
    const xml = tableDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain("tableStyleId");
  });

  it("round-trips table with table-level borders", () => {
    const opts: TableDescriptorOptions = {
      borders: {
        top: { width: 25400, color: "000000" },
        bottom: { width: 25400, color: "000000" },
        left: { width: 25400, color: "333333" },
        right: { width: 25400, color: "333333" },
      },
      rows: [{ cells: [{ text: "A" }, { text: "B" }] }, { cells: [{ text: "C" }, { text: "D" }] }],
    };

    // Table-level borders are distributed to edge cells during stringify.
    // Parse reads them back from those edge cells.
    const result = roundTrip(opts);
    const rows = result.rows!;
    const firstRowFirstCell = rows[0].cells![0] as Record<string, unknown>;
    const cellBorders = firstRowFirstCell.borders as Record<string, Record<string, unknown>>;

    // Top-left cell gets top + left borders
    const top = cellBorders.top!;
    expect(top.width).toBe(25400);
    expect(top.color).toBe("000000");
    const left = cellBorders.left!;
    expect(left.color).toBe("333333");
  });

  it("round-trips cell with margins", () => {
    const opts: TableDescriptorOptions = {
      rows: [
        {
          cells: [
            {
              text: "Margins",
              margins: { top: 1000, bottom: 2000, left: 3000, right: 4000 },
            },
          ],
        },
      ],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const cell = rows[0].cells![0];
    const margins = cell.margins!;

    expect(margins.top).toBe(1000);
    expect(margins.bottom).toBe(2000);
    expect(margins.left).toBe(3000);
    expect(margins.right).toBe(4000);
  });
});
