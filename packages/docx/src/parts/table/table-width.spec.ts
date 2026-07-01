import { describe, expect, it } from "vite-plus/test";

import { widthFiftiethsToPct, widthPctToFiftieths } from "./table-width";

describe("widthPctToFiftieths", () => {
  it("converts a percentage number to fiftieths (100 → 5000)", () => {
    expect(widthPctToFiftieths(100)).toBe(5000);
    expect(widthPctToFiftieths(50)).toBe(2500);
    expect(widthPctToFiftieths(0)).toBe(0);
  });

  it("converts a percentage string to fiftieths", () => {
    expect(widthPctToFiftieths("100%")).toBe(5000);
    expect(widthPctToFiftieths("50%")).toBe(2500);
  });

  it("passes UniversalMeasure through unchanged", () => {
    expect(widthPctToFiftieths("5cm")).toBe("5cm");
    expect(widthPctToFiftieths("2in")).toBe("2in");
  });

  it("handles negative percentages", () => {
    expect(widthPctToFiftieths(-100)).toBe(-5000);
  });
});

describe("widthFiftiethsToPct", () => {
  it("converts fiftieths to a percentage when type is pct", () => {
    expect(widthFiftiethsToPct(5000, "pct")).toBe(100);
    expect(widthFiftiethsToPct(2500, "pct")).toBe(50);
    expect(widthFiftiethsToPct(0, "pct")).toBe(0);
  });

  it("returns the raw size when type is not pct", () => {
    expect(widthFiftiethsToPct(5000, "dxa")).toBe(5000);
    expect(widthFiftiethsToPct(5000, undefined)).toBe(5000);
    // A non-number size (e.g. UniversalMeasure) is passed through even for pct.
    expect(widthFiftiethsToPct("5cm", "pct")).toBe("5cm");
  });

  it("handles negative fiftieths", () => {
    expect(widthFiftiethsToPct(-5000, "pct")).toBe(-100);
  });
});
