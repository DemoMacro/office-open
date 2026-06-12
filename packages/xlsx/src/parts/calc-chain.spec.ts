import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { calcChainDesc } from "./calc-chain";
import type { CalcChainOptions } from "./calc-chain";

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

function roundTrip(opts: CalcChainOptions) {
  const xml = calcChainDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return calcChainDesc.parse(el, readCtx) as unknown as CalcChainOptions;
}

// ── Tests ──

describe("calcChainDesc round-trip", () => {
  it("round-trips single cell", () => {
    const opts: CalcChainOptions = {
      cells: [{ reference: "A1", sheetIndex: 1 }],
    };
    const result = roundTrip(opts);

    expect(result.cells).toHaveLength(1);
    expect(result.cells![0].reference).toBe("A1");
    expect(result.cells![0].sheetIndex).toBe(1);
    expect(result.cells![0].array).toBeUndefined();
  });

  it("round-trips cell with array formula flag", () => {
    const opts: CalcChainOptions = {
      cells: [{ reference: "C3", sheetIndex: 2, array: true }],
    };
    const result = roundTrip(opts);

    expect(result.cells![0].array).toBe(true);
  });

  it("round-trips multiple cells across sheets", () => {
    const opts: CalcChainOptions = {
      cells: [
        { reference: "A1", sheetIndex: 1 },
        { reference: "B2", sheetIndex: 1 },
        { reference: "A1", sheetIndex: 2, array: true },
      ],
    };
    const result = roundTrip(opts);

    expect(result.cells).toHaveLength(3);
    expect(result.cells![0].reference).toBe("A1");
    expect(result.cells![0].sheetIndex).toBe(1);
    expect(result.cells![2].reference).toBe("A1");
    expect(result.cells![2].sheetIndex).toBe(2);
    expect(result.cells![2].array).toBe(true);
  });
});
