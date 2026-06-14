import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../context";
import { setTableParseChild, tableDesc } from "./descriptor";
import type { TableOptions } from "./table";

// Register a minimal child parser for table cells
setTableParseChild((_el, _ctx) => {
  return { paragraph: {} };
});

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) =>
    typeof child === "string" ? `<w:p><w:r><w:t>${child}</w:t></w:r></w:p>` : "<w:p/>",
  fileData: {} as never,
} as unknown as BodyContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: TableOptions) {
  const xml = tableDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return tableDesc.parse(el, readCtx);
}

describe("tableDesc round-trip", () => {
  it("round-trips a simple table", () => {
    const result = roundTrip({
      rows: [{ cells: [{ children: [{ paragraph: "cell1" }] }] }],
    });
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0].cells).toHaveLength(1);
  });

  it("round-trips table style", () => {
    const result = roundTrip({
      style: "TableGrid",
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.style).toBe("TableGrid");
  });

  it("round-trips table alignment", () => {
    const result = roundTrip({
      alignment: "center",
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.alignment).toBe("center");
  });

  it("round-trips table width", () => {
    const result = roundTrip({
      width: { size: 5000, type: "dxa" },
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.width).toBeDefined();
    expect(result.width!.size).toBe(5000);
  });

  it("round-trips table borders", () => {
    const result = roundTrip({
      borders: {
        top: { style: "single", color: "FF0000", size: 4 },
        bottom: { style: "single", color: "00FF00", size: 4 },
      },
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.borders).toBeDefined();
    expect(result.borders!.top!.style).toBe("single");
    expect(result.borders!.top!.color).toBe("FF0000");
    expect(result.borders!.top!.size).toBe(4);
  });

  it("does not inflate unspecified border sides", () => {
    // CT_TblBorders sides are all optional — setting only top must not invent
    // the other five sides (faithful round-trip: no inflation).
    const result = roundTrip({
      borders: { top: { style: "single" } },
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.borders!.top).toBeDefined();
    expect(result.borders!.bottom).toBeUndefined();
    expect(result.borders!.left).toBeUndefined();
    expect(result.borders!.insideHorizontal).toBeUndefined();
    expect(result.borders!.insideVertical).toBeUndefined();
  });

  it("round-trips inside borders", () => {
    // insideH/insideV XML names must map to insideHorizontal/insideVertical keys.
    const result = roundTrip({
      borders: {
        insideHorizontal: { style: "single", color: "0000FF" },
        insideVertical: { style: "single", color: "00FF00" },
      },
      rows: [
        { cells: [{ children: [] }, { children: [] }] },
        { cells: [{ children: [] }, { children: [] }] },
      ],
    });
    expect(result.borders!.insideHorizontal).toBeDefined();
    expect(result.borders!.insideHorizontal!.color).toBe("0000FF");
    expect(result.borders!.insideVertical).toBeDefined();
    expect(result.borders!.insideVertical!.color).toBe("00FF00");
  });

  it("round-trips border themeColor and shadow", () => {
    const result = roundTrip({
      borders: { top: { style: "single", themeColor: "text1", shadow: true } },
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.borders!.top!.themeColor).toBe("text1");
    expect(result.borders!.top!.shadow).toBe(true);
  });

  it("round-trips table layout", () => {
    const result = roundTrip({
      layout: "fixed",
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.layout).toBe("fixed");
  });

  it("round-trips column widths", () => {
    const result = roundTrip({
      columnWidths: [3000, 3000],
      rows: [{ cells: [{ children: [] }, { children: [] }] }],
    });
    expect(result.columnWidths).toEqual([3000, 3000]);
  });

  it("round-trips description", () => {
    const result = roundTrip({
      description: "Test table description",
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.description).toBe("Test table description");
  });

  it("round-trips multiple rows", () => {
    const result = roundTrip({
      rows: [
        { cells: [{ children: [{ paragraph: "r1c1" }] }] },
        { cells: [{ children: [{ paragraph: "r2c1" }] }] },
        { cells: [{ children: [{ paragraph: "r3c1" }] }] },
      ],
    });
    expect(result.rows).toHaveLength(3);
  });

  it("round-trips row with height", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [{ children: [] }],
          height: { value: 500, rule: "atLeast" },
        },
      ],
    });
    expect(result.rows[0].height).toBeDefined();
    expect(result.rows[0].height!.value).toBe(500);
  });

  it("round-trips cell with column span", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [{ children: [], columnSpan: 2 }, { children: [] }],
        },
      ],
    });
    expect(result.rows[0].cells[0].columnSpan).toBe(2);
  });

  it("round-trips cell shading", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            {
              children: [],
              shading: { fill: "FFFF00", type: "clear" },
            },
          ],
        },
      ],
    });
    const cell = result.rows[0].cells[0];
    expect(cell.shading).toBeDefined();
    expect(cell.shading!.fill).toBe("FFFF00");
  });

  it("round-trips multiple cells in a row", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            { children: [{ paragraph: "a" }] },
            { children: [{ paragraph: "b" }] },
            { children: [{ paragraph: "c" }] },
          ],
        },
      ],
    });
    expect(result.rows[0].cells).toHaveLength(3);
  });
});
