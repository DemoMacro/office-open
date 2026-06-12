import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { tableDesc } from "./table";
import type { TableOptions } from "./table";

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
  const doc = parseXml(xml);
  const el = doc.elements![0];
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
    const columns = result.columns!;
    expect(columns).toHaveLength(2);
    expect(columns[0].name).toBe("Col1");
    expect(columns[1].name).toBe("Col2");
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

    expect(columns[1].totalsRowFunction).toBe("sum");
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

    expect(columns[2].calculatedColumnFormula).toBe("A+B");
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

    expect(columns[0].totalsRowLabel).toBe("Total");
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
});
