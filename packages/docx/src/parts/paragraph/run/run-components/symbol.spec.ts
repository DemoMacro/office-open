import { describe, expect, it } from "vite-plus/test";

import { buildSymbol } from "./symbol";

describe("buildSymbol", () => {
  it("creates an empty symbol with default font", () => {
    expect(buildSymbol()).toEqual({
      "w:sym": { _attr: { "w:char": "", "w:font": "Wingdings" } },
    });
  });

  it("creates the provided symbol with default font", () => {
    expect(buildSymbol("F071")).toEqual({
      "w:sym": { _attr: { "w:char": "F071", "w:font": "Wingdings" } },
    });
  });

  it("creates the provided symbol with the provided font", () => {
    expect(buildSymbol("F071", "Arial")).toEqual({
      "w:sym": { _attr: { "w:char": "F071", "w:font": "Arial" } },
    });
  });
});
