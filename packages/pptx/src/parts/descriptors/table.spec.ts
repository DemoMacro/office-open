import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { TableOptions } from "@shared/table/table-frame";
import { describe, expect, it } from "vite-plus/test";

import { tableDesc } from "./table";

// ── Mock PPTX write context ──

class MockWriteContext {
  registerShapeId() {}
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

function roundTrip(opts: TableOptions) {
  const writeCtx = new MockWriteContext() as unknown as WriteContext;
  const xml = tableDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return tableDesc.parse(el, readCtx);
}

describe("tableDesc round-trip", () => {
  it("round-trips basic 2x2 table", () => {
    const opts: TableOptions = {
      rows: [
        { cells: [{ text: "A1" }, { text: "B1" }] },
        { cells: [{ text: "A2" }, { text: "B2" }] },
      ],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;

    expect(rows).toHaveLength(2);
    const [row0] = rows;
    const row0Cells = row0!.cells!;
    expect(row0Cells).toHaveLength(2);
    expect(row0Cells[0]?.text).toBe("A1");
    expect(row0Cells[1]?.text).toBe("B1");
  });

  it("round-trips table with position", () => {
    const opts: TableOptions = {
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
    const opts: TableOptions = {
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
    const opts: TableOptions = {
      rows: [{ height: 500, cells: [{ text: "Tall" }] }],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const [row0] = rows;

    expect(row0?.height).toBe(500);
  });

  it("round-trips cell with fill", () => {
    const opts: TableOptions = {
      rows: [
        {
          cells: [{ text: "Filled", fill: { type: "solid", color: "4472C4" } }],
        },
      ],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const cell = rows[0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");
    const fill = cell.fill! as { type: string; color: { value: string } };

    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("4472C4");
  });

  it("round-trips cell with vert (text direction)", () => {
    const opts: TableOptions = {
      rows: [
        {
          cells: [{ text: "Rotated", vertical: "vertical270" }],
        },
      ],
    };
    const result = roundTrip(opts);
    const cell = result.rows![0]?.cells?.[0];

    expect(cell?.vertical).toBe("vertical270");
  });

  it("round-trips cell with borders", () => {
    const opts: TableOptions = {
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
    const cell = rows[0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");
    const borders = cell.borders!;
    const top = borders.top!;

    expect(top.width).toBe(12700);
    expect(top.color).toBe("000000");
  });

  it("round-trips a scheme-color border as structured fill", () => {
    const opts: TableOptions = {
      rows: [
        {
          cells: [
            {
              text: "Scheme",
              borders: {
                top: { width: 12700, color: { type: "solid", color: { value: "accent1" } } },
              },
            },
          ],
        },
      ],
    };
    const result = roundTrip(opts);
    const cell = result.rows![0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");
    expect(cell.borders?.top?.color).toEqual({ type: "solid", color: { value: "accent1" } });
  });

  it("round-trips cell with diagonal borders", () => {
    const opts: TableOptions = {
      rows: [
        {
          cells: [
            {
              text: "Crossed",
              borders: {
                diagonalTopLeftToBottomRight: { width: 12700, color: "C00000" },
                diagonalBottomLeftToTopRight: { width: 9525, color: "000000" },
              },
            },
          ],
        },
      ],
    };
    const result = roundTrip(opts);
    const cell = result.rows![0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");
    expect(cell.borders?.diagonalTopLeftToBottomRight).toMatchObject({
      width: 12700,
      color: "C00000",
    });
    expect(cell.borders?.diagonalBottomLeftToTopRight).toMatchObject({
      width: 9525,
      color: "000000",
    });
  });

  it("round-trips cell with verticalAlign", () => {
    const opts: TableOptions = {
      rows: [{ cells: [{ text: "Center", verticalAlign: "center" }] }],
    };
    const result = roundTrip(opts);
    const rows = result.rows!;
    const cell = rows[0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");

    expect(cell.verticalAlign).toBe("center");
  });

  it("writes multiple cell properties as tcPr attributes", () => {
    const writeCtx = new MockWriteContext() as unknown as WriteContext;
    const xml = tableDesc.stringify(
      {
        rows: [
          {
            cells: [
              {
                text: "Cell",
                verticalAlign: "center",
                vertical: "vertical270",
                margins: { left: "2.5mm", right: 365760, top: 274320, bottom: 457200 },
              },
            ],
          },
        ],
      },
      writeCtx,
    )!;
    expect(xml).toContain(
      '<a:tcPr anchor="ctr" vert="vert270" marL="90000" marR="365760" marT="274320" marB="457200"/>',
    );
    expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns)=/);
  });

  it("round-trips table properties (bandRow, firstRow, etc.)", () => {
    const opts: TableOptions = {
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
    const opts: TableOptions = {
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
    const opts: TableOptions = {
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
    const firstRowFirstCell = (rows[0]?.cells?.[0] ?? undefined) as Record<string, unknown>;
    const cellBorders = firstRowFirstCell.borders as Record<string, Record<string, unknown>>;

    // Top-left cell gets top + left borders
    const top = cellBorders.top!;
    expect(top.width).toBe(25400);
    expect(top.color).toBe("000000");
    const left = cellBorders.left!;
    expect(left.color).toBe("333333");
  });

  it("round-trips cell with margins", () => {
    const opts: TableOptions = {
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
    const cell = rows[0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");
    const margins = cell.margins!;

    expect(margins.top).toBe(1000);
    expect(margins.bottom).toBe(2000);
    expect(margins.left).toBe(3000);
    expect(margins.right).toBe(4000);
  });

  it("parses source tcPr margin attributes and writes them back", () => {
    const source =
      '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGraphicFramePr><p:cNvPr id="1" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000"/></a:tblGrid><a:tr h="1000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr marL="182880" marR="365760" marT="274320" marB="457200"/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
    const root = parseXml(source).elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    const parsed = tableDesc.parse(root, readCtx);
    const writeCtx = new MockWriteContext() as unknown as WriteContext;
    const xml = tableDesc.stringify(parsed, writeCtx)!;

    expect(parsed.rows?.[0]?.cells?.[0]?.margins).toEqual({
      left: 182880,
      right: 365760,
      top: 274320,
      bottom: 457200,
    });
    expect(xml).toContain('<a:tcPr marL="182880" marR="365760" marT="274320" marB="457200"/>');
  });

  it("round-trips cell margins with UniversalMeasure (mm)", () => {
    const opts: TableOptions = {
      rows: [
        {
          cells: [
            {
              text: "M",
              margins: { top: "1mm", left: "2.5mm" },
            },
          ],
        },
      ],
    };
    const result = roundTrip(opts);
    const cell = result.rows![0]?.cells?.[0];
    if (!cell) throw new Error("missing cell");
    const margins = cell.margins!;
    // UniversalMeasure input is normalized to EMU on stringify (ST_TextMargin
    // is an integer type), so it comes back as the plain EMU number.
    expect(margins.top).toBe(36000);
    expect(margins.left).toBe(90000);
  });

  it("round-trips cell 3D bevel (a:cell3D)", () => {
    const opts: TableOptions = {
      rows: [
        {
          cells: [
            {
              text: "A1",
              cell3D: {
                prstMaterial: "metal",
                bevel: { w: 25400, h: 19050, prst: "circle" },
                lightRig: { rig: "threePt", direction: "top" },
              },
            },
          ],
        },
      ],
    };
    const result = roundTrip(opts);
    expect(result.rows[0]!.cells[0]!.cell3D).toMatchObject({
      prstMaterial: "metal",
      bevel: { w: 25400, h: 19050, prst: "circle" },
      lightRig: { rig: "threePt", direction: "top" },
    });
  });

  it("round-trips inline table style (a:tableStyle in a:tblPr)", () => {
    const opts: TableOptions = {
      rows: [{ cells: [{ text: "A1" }] }],
      tableStyle: {
        styleId: "{5940675A-B579-460E-94D1-54222C63F5DA}",
        styleName: "Inline Style",
        regions: { wholeTbl: { text: { bold: "on" } } },
      },
    };
    const result = roundTrip(opts);
    expect(result.tableStyle?.styleId).toBe("{5940675A-B579-460E-94D1-54222C63F5DA}");
    expect(result.tableStyle?.styleName).toBe("Inline Style");
    expect(result.tableStyle?.regions?.wholeTbl?.text?.bold).toBe("on");
  });

  it("round-trips tableStyleId (a:tableStyleId)", () => {
    const opts: TableOptions = {
      rows: [{ cells: [{ text: "A1" }] }],
      tableStyleId: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    };
    const result = roundTrip(opts);

    expect(result.tableStyleId).toBe("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
  });
});
