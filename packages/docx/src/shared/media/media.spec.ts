import { describe, expect, it } from "vite-plus/test";

import { createTransformation } from "./media";

describe("createTransformation", () => {
  it("treats numbers as native EMUs", () => {
    const result = createTransformation({ width: 952500, height: 1905000 });

    expect(result).to.deep.equal({
      width: 952500,
      height: 1905000,
      flipHorizontal: undefined,
      flipVertical: undefined,
      rotation: undefined,
    });
  });

  it("converts universal measures to EMU", () => {
    const result = createTransformation({ width: "100px", height: "200px" });

    expect(result.width).toBe(952500);
    expect(result.height).toBe(1905000);
  });

  it("converts offset UniversalMeasure to EMU", () => {
    const result = createTransformation({
      height: 50,
      offset: { x: "10px", y: "20px" },
      width: 100,
    });

    expect(result.offset).to.deep.equal({ x: 95250, y: 190500 });
  });

  it("should pass rotation through in degrees", () => {
    const result = createTransformation({
      height: 100,
      rotation: 90,
      width: 100,
    });

    expect(result.rotation).to.equal(90);
  });

  it("should pass flips through unchanged", () => {
    const result = createTransformation({
      flipHorizontal: true,
      height: 100,
      width: 100,
    });

    expect(result.flipHorizontal).to.equal(true);
    expect(result.flipVertical).to.equal(undefined);
  });
});
