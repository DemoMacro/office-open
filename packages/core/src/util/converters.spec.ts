import { describe, expect, it } from "vite-plus/test";

import {
  convertInchesToTwip,
  convertMillimetersToTwip,
  convertToEmu,
  convertUniversalMeasureToEmu,
  convertUniversalMeasureToTwip,
  parseUniversalMeasure,
  emitPercent,
  parsePercent,
  emitAngle,
  parseAngle,
} from "./converters";

describe("convertMillimetersToTwip", () => {
  it("should convert 25.4mm to 1440 twips (1 inch)", () => {
    expect(convertMillimetersToTwip(25.4)).toBe(1440);
  });

  it("should convert 1mm correctly", () => {
    // 1mm = 1/25.4 * 72 * 20 ≈ 56.69
    expect(convertMillimetersToTwip(1)).toBe(Math.floor((1 / 25.4) * 72 * 20));
  });

  it("should handle zero", () => {
    expect(convertMillimetersToTwip(0)).toBe(0);
  });

  it("should handle small values", () => {
    expect(convertMillimetersToTwip(0.1)).toBe(Math.floor((0.1 / 25.4) * 72 * 20));
  });
});

describe("convertInchesToTwip", () => {
  it("should convert 1 inch to 1440 twips", () => {
    expect(convertInchesToTwip(1)).toBe(1440);
  });

  it("should convert 0.5 inch to 720 twips", () => {
    expect(convertInchesToTwip(0.5)).toBe(720);
  });

  it("should handle zero", () => {
    expect(convertInchesToTwip(0)).toBe(0);
  });

  it("should be consistent with mm conversion", () => {
    // 1 inch = 25.4mm
    expect(convertInchesToTwip(1)).toBe(convertMillimetersToTwip(25.4));
  });
});

describe("px unit (project extension, 96 DPI)", () => {
  it("convertToEmu parses px via UniversalMeasure, number stays EMU", () => {
    expect(convertToEmu("200px")).toBe(200 * 9525); // 1905000
    expect(convertToEmu(914400)).toBe(914400); // number is already EMU
    expect(convertToEmu("1in")).toBe(914400); // other units still work
  });

  it("convertUniversalMeasureToEmu handles px (1px = 9525 EMU)", () => {
    expect(convertUniversalMeasureToEmu("200px")).toBe(1905000);
    expect(convertUniversalMeasureToEmu("1in")).toBe(914400);
  });

  it("convertUniversalMeasureToTwip handles px (1px = 15twip)", () => {
    expect(convertUniversalMeasureToTwip("200px")).toBe(3000);
    expect(convertUniversalMeasureToTwip("1in")).toBe(1440);
  });

  it("parseUniversalMeasure recognizes px", () => {
    expect(parseUniversalMeasure("200px")).toEqual({ value: 200, unit: "px" });
  });
});

describe("emitPercent", () => {
  it("should convert integer percent to scalar", () => {
    expect(emitPercent(50)).toBe(50000);
    expect(emitPercent(100)).toBe(100000);
  });

  it("should round fractional percent", () => {
    expect(emitPercent(12.3456)).toBe(12346);
    expect(emitPercent(99.9999)).toBe(100000);
  });

  it("should handle zero", () => {
    expect(emitPercent(0)).toBe(0);
  });

  it("should handle negative percent", () => {
    expect(emitPercent(-50)).toBe(-50000);
    expect(emitPercent(-12.5)).toBe(-12500);
  });
});

describe("parsePercent", () => {
  it("should convert scalar to percent", () => {
    expect(parsePercent(50000)).toBe(50);
    expect(parsePercent(100000)).toBe(100);
  });

  it("should preserve fractional precision", () => {
    expect(parsePercent(12345)).toBe(12.345);
  });

  it("should handle zero", () => {
    expect(parsePercent(0)).toBe(0);
  });

  it("should handle negative scalar", () => {
    expect(parsePercent(-50000)).toBe(-50);
  });
});

describe("emitAngle", () => {
  it("should convert integer degrees to scalar", () => {
    expect(emitAngle(90)).toBe(5400000);
    expect(emitAngle(180)).toBe(10800000);
  });

  it("should round fractional degrees", () => {
    expect(emitAngle(1.23456)).toBe(74074);
    expect(emitAngle(-1.23456)).toBe(-74074);
  });

  it("should handle zero", () => {
    expect(emitAngle(0)).toBe(0);
  });

  it("should handle negative angles", () => {
    expect(emitAngle(-90)).toBe(-5400000);
  });
});

describe("parseAngle", () => {
  it("should convert scalar to degrees", () => {
    expect(parseAngle(5400000)).toBe(90);
    expect(parseAngle(10800000)).toBe(180);
  });

  it("should preserve fractional precision", () => {
    expect(parseAngle(74074)).toBeCloseTo(1.23456666, 5);
  });

  it("should handle zero", () => {
    expect(parseAngle(0)).toBe(0);
  });

  it("should handle negative scalar", () => {
    expect(parseAngle(-5400000)).toBe(-90);
  });
});
