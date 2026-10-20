import { describe, expect, it } from "vitest";

import { decodeFormula } from "./formula";

function uint16(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, value >> 8]);
}

function bytes(...values: readonly number[]): Uint8Array {
  return new Uint8Array(values);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((total, part) => total + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

describe("decodeFormula", () => {
  it("decodes error and string constants", () => {
    expect(decodeFormula(new Uint8Array([0x1c, 0x0f]), { baseRow: 0, baseColumn: 0 }).formula).toBe(
      "#VALUE!",
    );
    expect(
      decodeFormula(concat([bytes(0x17), uint16(2), bytes(0), bytes(0x61, 0x62)]), {
        baseRow: 0,
        baseColumn: 0,
      }).formula,
    ).toBe('"ab"');
  });

  it("decodes memory functions and deleted area references", () => {
    expect(
      decodeFormula(concat([bytes(0x29), uint16(3), bytes(0x1e), uint16(42)]), {
        baseRow: 0,
        baseColumn: 0,
      }).formula,
    ).toBe("42");
    expect(
      decodeFormula(concat([bytes(0x3d), uint16(0), new Uint8Array(8)]), {
        baseRow: 0,
        baseColumn: 0,
        sheetNameByExternIndex: () => "Sheet2",
      }).formula,
    ).toBe("Sheet2!#REF!");
  });

  it("decodes variable function tokens in argument-first order", () => {
    expect(
      decodeFormula(Uint8Array.from([0x44, 0, 0, 0, 0x40, 0x42, 1, 4, 0]), {
        baseRow: 0,
        baseColumn: 0,
      }).formula,
    ).toBe("SUM($A$1)");
  });

  it("decodes array constants from the token-array trailer", () => {
    const tokens = concat([bytes(0x60), new Uint8Array(7)]);
    const arrayData = concat([
      new Uint8Array([0, 0]),
      new Uint8Array([1]),
      new Uint8Array([0, 0, 0, 0, 0, 0, 0xf0, 0x3f]),
    ]);
    expect(decodeFormula(tokens, { baseRow: 0, baseColumn: 0, arrayData }).formula).toBe("{1}");
  });
});
