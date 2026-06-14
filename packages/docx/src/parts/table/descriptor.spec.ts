import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../context";
import { setTableParseChild, tableDesc } from "./descriptor";
import type { TableOptions } from "./table";
import type { TableCellOptions } from "./table-cell";
import type { TableRowOptions } from "./table-row";

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
    expect((result.rows[0] as TableRowOptions).cells).toHaveLength(1);
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
    const row = result.rows[0] as TableRowOptions;
    expect(row.height).toBeDefined();
    expect(row.height!.value).toBe(500);
  });

  it("round-trips cell with column span", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [{ children: [], columnSpan: 2 }, { children: [] }],
        },
      ],
    });
    expect(((result.rows[0] as TableRowOptions).cells[0] as TableCellOptions).columnSpan).toBe(2);
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
    const cell = (result.rows[0] as TableRowOptions).cells[0] as TableCellOptions;
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
    expect((result.rows[0] as TableRowOptions).cells).toHaveLength(3);
  });

  it("round-trips visuallyRightToLeft (bidiVisual)", () => {
    const result = roundTrip({
      visuallyRightToLeft: true,
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.visuallyRightToLeft).toBe(true);
  });

  it("round-trips styleRowBandSize / styleColBandSize", () => {
    const result = roundTrip({
      styleRowBandSize: 2,
      styleColBandSize: 3,
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.styleRowBandSize).toBe(2);
    expect(result.styleColBandSize).toBe(3);
  });

  it("round-trips caption (tblCaption)", () => {
    const result = roundTrip({
      caption: "My Caption",
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.caption).toBe("My Caption");
  });

  it("round-trips cellSpacing (tblCellSpacing)", () => {
    const result = roundTrip({
      cellSpacing: { value: 108, type: "dxa" },
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.cellSpacing).toBeDefined();
    expect(result.cellSpacing!.value).toBe(108);
  });

  it("round-trips tblPrChange revision", () => {
    const result = roundTrip({
      revision: { id: 1, author: "A", date: "2024-01-01T00:00:00Z", style: "TableGrid" },
      rows: [{ cells: [{ children: [] }] }],
    });
    expect(result.revision).toBeDefined();
    expect(result.revision!.id).toBe(1);
    expect(result.revision!.author).toBe("A");
  });

  it("round-trips row rsid attributes", () => {
    const result = roundTrip({
      rows: [
        {
          rsidRPr: "00112233",
          rsidR: "00AABBCC",
          rsidDel: "00DDEEFF",
          rsidTr: "00445566",
          cells: [{ children: [] }],
        },
      ],
    });
    const row = result.rows[0] as TableRowOptions;
    expect(row.rsidRPr).toBe("00112233");
    expect(row.rsidR).toBe("00AABBCC");
    expect(row.rsidDel).toBe("00DDEEFF");
    expect(row.rsidTr).toBe("00445566");
  });

  it("round-trips row trPr fields (cnfStyle/divId/grid/gridBefore/gridAfter/wBefore/wAfter/jc/hidden)", () => {
    const result = roundTrip({
      rows: [
        {
          cnfStyle: { val: "000000010000" },
          divId: 5,
          gridBefore: 1,
          gridAfter: 2,
          widthBefore: { size: 100, type: "dxa" },
          widthAfter: { size: 200, type: "dxa" },
          rowAlignment: "center",
          hidden: true,
          cells: [{ children: [] }],
        },
      ],
    });
    const row = result.rows[0] as TableRowOptions;
    expect(row.cnfStyle?.val).toBe("000000010000");
    expect(row.divId).toBe(5);
    expect(row.gridBefore).toBe(1);
    expect(row.gridAfter).toBe(2);
    expect(row.widthBefore?.size).toBe(100);
    expect(row.widthAfter?.size).toBe(200);
    expect(row.rowAlignment).toBe("center");
    expect(row.hidden).toBe(true);
  });

  it("round-trips row trPrChange revision", () => {
    const result = roundTrip({
      rows: [
        {
          revision: { id: 2, author: "B", date: "2024-02-02T00:00:00Z", hidden: true },
          cells: [{ children: [] }],
        },
      ],
    });
    const row = result.rows[0] as TableRowOptions;
    expect(row.revision).toBeDefined();
    expect(row.revision!.id).toBe(2);
    expect(row.revision!.hidden).toBe(true);
  });

  it("round-trips cell diagonal borders (tl2br/tr2bl) and start/end", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            {
              children: [],
              borders: {
                start: { style: "single", color: "FF0000" },
                end: { style: "single", color: "00FF00" },
                topLeftToBottomRight: { style: "single", color: "0000FF" },
                topRightToBottomLeft: { style: "single", color: "FFFF00" },
              },
            },
          ],
        },
      ],
    });
    const cell = (result.rows[0] as TableRowOptions).cells[0] as TableCellOptions;
    expect(cell.borders!.start!.color).toBe("FF0000");
    expect(cell.borders!.end!.color).toBe("00FF00");
    expect(cell.borders!.topLeftToBottomRight!.color).toBe("0000FF");
    expect(cell.borders!.topRightToBottomLeft!.color).toBe("FFFF00");
  });

  it("round-trips cell hMerge/tcFitText/hideMark/headers", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            {
              children: [],
              horizontalMerge: "restart",
              fitText: true,
              hideMark: true,
              headers: ["h1", "h2"],
            },
          ],
        },
      ],
    });
    const cell = (result.rows[0] as TableRowOptions).cells[0] as TableCellOptions;
    expect(cell.horizontalMerge).toBe("restart");
    expect(cell.fitText).toBe(true);
    expect(cell.hideMark).toBe(true);
    expect(cell.headers).toEqual(["h1", "h2"]);
  });

  it("round-trips cell cellIns/cellDel and tcPrChange revision", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            {
              children: [],
              insertion: { id: 1, author: "A", date: "2024-01-01T00:00:00Z" },
              deletion: { id: 2, author: "B", date: "2024-02-02T00:00:00Z" },
              revision: { id: 3, author: "C", date: "2024-03-03T00:00:00Z", hideMark: true },
            },
          ],
        },
      ],
    });
    const cell = (result.rows[0] as TableRowOptions).cells[0] as TableCellOptions;
    expect(cell.insertion?.id).toBe(1);
    expect(cell.insertion?.author).toBe("A");
    expect(cell.deletion?.id).toBe(2);
    expect(cell.revision?.id).toBe(3);
    expect(cell.revision?.hideMark).toBe(true);
  });

  it("round-trips cell cellMerge (vMerge/vMergeOrig)", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            {
              children: [],
              cellMerge: {
                id: 4,
                author: "D",
                date: "2024-04-04T00:00:00Z",
                verticalMerge: "restart",
                verticalMergeOriginal: "continue",
              },
            },
          ],
        },
      ],
    });
    const cell = (result.rows[0] as TableRowOptions).cells[0] as TableCellOptions;
    expect(cell.cellMerge?.id).toBe(4);
    expect(cell.cellMerge?.verticalMerge).toBe("restart");
    expect(cell.cellMerge?.verticalMergeOriginal).toBe("continue");
  });

  it("round-trips a cell-level SDT (CT_SdtCell wrapping a cell)", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            { children: [{ paragraph: "normal" }] },
            {
              sdt: {
                properties: { alias: "CellCtrl", tag: "cell-ctrl", text: {} },
                cells: [{ children: [{ paragraph: "controlled" }] }],
              },
            },
          ],
        },
      ],
    });
    const cells = (result.rows[0] as TableRowOptions).cells;
    expect(cells).toHaveLength(2);
    expect("sdt" in cells[1]).toBe(true);
    if ("sdt" in cells[1]) {
      expect(cells[1].sdt.properties).toMatchObject({ alias: "CellCtrl", tag: "cell-ctrl" });
      expect(cells[1].sdt.cells).toHaveLength(1);
    }
  });

  it("round-trips a row-level SDT (CT_SdtRow wrapping a row)", () => {
    const result = roundTrip({
      rows: [
        { cells: [{ children: [{ paragraph: "header" }] }] },
        {
          sdt: {
            properties: { alias: "RowCtrl", tag: "row-ctrl", richText: true },
            rows: [{ cells: [{ children: [{ paragraph: "controlled row" }] }] }],
          },
        },
      ],
    });
    expect(result.rows).toHaveLength(2);
    expect("sdt" in result.rows[1]).toBe(true);
    if ("sdt" in result.rows[1]) {
      expect(result.rows[1].sdt.properties).toMatchObject({ alias: "RowCtrl", tag: "row-ctrl" });
      expect(result.rows[1].sdt.rows).toHaveLength(1);
    }
  });

  it("round-trips a cell-level customXml (CT_CustomXmlCell wrapping a cell)", () => {
    const result = roundTrip({
      rows: [
        {
          cells: [
            { children: [{ paragraph: "normal" }] },
            {
              customXml: {
                element: "taggedCell",
                uri: "http://ns",
                customXmlPr: { attributes: [{ name: "k", val: "v" }] },
                children: [{ children: [{ paragraph: "wrapped" }] }],
              },
            },
          ],
        },
      ],
    });
    const cells = (result.rows[0] as TableRowOptions).cells;
    expect(cells).toHaveLength(2);
    expect("customXml" in cells[1]).toBe(true);
    if ("customXml" in cells[1]) {
      expect(cells[1].customXml.element).toBe("taggedCell");
      expect(cells[1].customXml.uri).toBe("http://ns");
      expect(cells[1].customXml.customXmlPr?.attributes).toEqual([{ name: "k", val: "v" }]);
      expect(cells[1].customXml.children).toHaveLength(1);
    }
  });

  it("round-trips a row-level customXml (CT_CustomXmlRow wrapping a row)", () => {
    const result = roundTrip({
      rows: [
        { cells: [{ children: [{ paragraph: "header" }] }] },
        {
          customXml: {
            element: "taggedRow",
            children: [{ cells: [{ children: [{ paragraph: "wrapped row" }] }] }],
          },
        },
      ],
    });
    expect(result.rows).toHaveLength(2);
    expect("customXml" in result.rows[1]).toBe(true);
    if ("customXml" in result.rows[1]) {
      expect(result.rows[1].customXml.element).toBe("taggedRow");
      expect(result.rows[1].customXml.children).toHaveLength(1);
    }
  });
});
