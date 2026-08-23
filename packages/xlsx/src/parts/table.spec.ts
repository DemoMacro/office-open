import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { tableDesc } from "./table";
import type { TableOptions } from "./table";
import type { AutoFilterOptions } from "./worksheet";

// ── Minimal context stubs ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: TableOptions) {
  const xml = tableDesc.stringify(opts, writeCtx)!;
  // nativeTypeAttributes mirrors the real xlsx parse path (ParsedArchive.get
  // coerces "1"/"0" to numbers), so boolean attribute reads are exercised
  // against numeric coercion rather than a permissive non-coerced parse.
  const doc = parseXml(xml, { nativeTypeAttributes: true });
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return tableDesc.parse(el, readCtx);
}

// ── Tests ──

describe("tableDesc round-trip", () => {
  it("round-trips minimal table", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "Table1",
      ref: "A1:D4",
      columns: [{ name: "Col1" }, { name: "Col2" }],
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(1);
    expect(result.displayName).toBe("Table1");
    expect(result.ref).toBe("A1:D4");
    // Ref spans 4 columns — missing declarations pad with default ColumnN
    // names (Excel requires the count to match the ref width).
    const columns = result.columns!;
    expect(columns).toHaveLength(4);
    expect(columns[0]?.name).toBe("Col1");
    expect(columns[1]?.name).toBe("Col2");
    expect(columns[2]?.name).toBe("Column3");
    expect(columns[3]?.name).toBe("Column4");
  });

  it("round-trips table with name", () => {
    const opts: TableOptions = {
      id: 2,
      name: "MyTable",
      displayName: "MyTable",
      ref: "A1:C10",
      columns: [{ name: "A" }],
    };
    const result = roundTrip(opts);

    expect(result.name).toBe("MyTable");
    expect(result.displayName).toBe("MyTable");
  });

  it("round-trips table with autoFilter", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      autoFilter: "A1:B5",
      columns: [{ name: "X" }, { name: "Y" }],
    };
    const result = roundTrip(opts);

    expect(result.autoFilter).toBe("A1:B5");
  });

  it("round-trips column with totalsRowFunction", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      columns: [{ name: "Label" }, { name: "Value", totalsRowFunction: "sum" }],
    };
    const result = roundTrip(opts);
    const columns = result.columns!;

    expect(columns[1]?.totalsRowFunction).toBe("sum");
  });

  it("round-trips column with calculatedColumnFormula", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:C5",
      columns: [{ name: "A" }, { name: "B" }, { name: "C", calculatedColumnFormula: "A+B" }],
    };
    const result = roundTrip(opts);
    const columns = result.columns!;

    expect(columns[2]?.calculatedColumnFormula).toBe("A+B");
  });

  it("round-trips table with style", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      columns: [{ name: "X" }],
      style: {
        name: "TableStyleMedium2",
        showFirstColumn: true,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    };
    const result = roundTrip(opts);
    const style = result.style!;

    expect(style.name).toBe("TableStyleMedium2");
    expect(style.showFirstColumn).toBe(true);
    expect(style.showRowStripes).toBe(true);
  });

  it("round-trips styleless table as style=undefined", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);

    expect(result.style).toBeUndefined();
  });

  it("round-trips headerRowCount and totalsRowCount", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      headerRowCount: 0,
      totalsRowCount: 1,
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);

    expect(result.headerRowCount).toBe(0);
    expect(result.totalsRowCount).toBe(1);
  });

  it("round-trips insertRowShift and published flags", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      insertRowShift: true,
      published: true,
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);

    expect(result.insertRowShift).toBe(true);
    expect(result.published).toBe(true);
  });

  it("round-trips column with totalsRowLabel", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      columns: [{ name: "A", totalsRowFunction: "none", totalsRowLabel: "Total" }],
    };
    const result = roundTrip(opts);
    const columns = result.columns!;

    expect(columns[0]?.totalsRowLabel).toBe("Total");
  });

  it("round-trips totalsRowShown=false", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      totalsRowCount: 1,
      totalsRowShown: false,
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);

    expect(result.totalsRowShown).toBe(false);
  });

  it("round-trips comment, insertRow, connectionId attributes", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      comment: "Source table",
      ref: "A1:B5",
      insertRow: true,
      connectionId: 42,
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);

    expect(result.comment).toBe("Source table");
    expect(result.insertRow).toBe(true);
    expect(result.connectionId).toBe(42);
  });

  it("round-trips all tableStyleInfo boolean flags under nativeTypeAttributes coercion", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      columns: [{ name: "X" }],
      style: {
        name: "TableStyleMedium2",
        showFirstColumn: true,
        showLastColumn: true,
        showColumnStripes: true,
      },
    };
    const result = roundTrip(opts);
    const style = result.style!;

    // "1" coerces to number 1 on the real parse path; String() guard recovers it
    expect(style.showFirstColumn).toBe(true);
    expect(style.showLastColumn).toBe(true);
    expect(style.showColumnStripes).toBe(true);
  });

  it("round-trips ref-only autoFilter as shorthand string", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      autoFilter: "A1:B4",
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);

    expect(result.autoFilter).toBe("A1:B4");
  });

  it("round-trips structured autoFilter with filter columns and sort state", () => {
    const opts: TableOptions = {
      id: 1,
      displayName: "T1",
      ref: "A1:B5",
      autoFilter: {
        ref: "A1:B4",
        columns: [{ colId: 0, customFilters: { entries: [{ operator: "equal", val: "yes" }] } }],
        sortState: { ref: "A1:B4", conditions: [{ ref: "A1", descending: true }] },
      },
      columns: [{ name: "X" }],
    };
    const result = roundTrip(opts);
    const af = result.autoFilter as AutoFilterOptions;

    expect(af.ref).toBe("A1:B4");
    expect(af.columns).toEqual([
      { colId: 0, customFilters: { entries: [{ operator: "equal", val: "yes" }] } },
    ]);
    expect(af.sortState).toEqual({
      ref: "A1:B4",
      conditions: [{ ref: "A1", descending: true }],
    });
  });

  it("round-trips the trailing extLst verbatim", () => {
    const ext =
      '<ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
      '<x14:table altTextSummary="Cost summary"/></ext>';
    const opts: TableOptions = {
      id: 1,
      displayName: "Table1",
      ref: "A1:D4",
      columns: [{ name: "Col1" }],
      ext,
    };
    const xml = tableDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain("<extLst>" + ext + "</extLst>");

    const result = roundTrip(opts);
    expect(result.ext).toBe(ext);
  });
});
