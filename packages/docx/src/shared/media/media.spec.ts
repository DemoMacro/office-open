import { describe, expect, it } from "vite-plus/test";

import { createTransformation } from "./media";

describe("createTransformation", () => {
  it("treats numbers as EMU and derives pixels", () => {
    // number = EMU (native, lossless); pixels derived via 1px = 9525 EMU
    const result = createTransformation({ width: 952500, height: 1905000 });

    expect(result).to.deep.equal({
      emus: {
        x: 952500,
        y: 1905000,
      },
      flip: undefined,
      offset: {
        emus: {
          x: 0,
          y: 0,
        },
        pixels: {
          x: 0,
          y: 0,
        },
      },
      pixels: {
        x: 100,
        y: 200,
      },
      rotation: undefined,
    });
  });

  it("converts UniversalMeasure (px) to EMU", () => {
    const result = createTransformation({ width: "100px", height: "200px" });

    expect(result.emus).to.deep.equal({ x: 952500, y: 1905000 });
    expect(result.pixels).to.deep.equal({ x: 100, y: 200 });
  });

  it("converts offset UniversalMeasure to EMU", () => {
    const result = createTransformation({
      height: 50,
      offset: { left: "10px", top: "20px" },
      width: 100,
    });

    expect(result.offset).to.deep.equal({
      emus: { x: 95250, y: 190500 },
      pixels: { x: 10, y: 20 },
    });
  });

  it("should convert rotation from degrees to 60000ths", () => {
    const result = createTransformation({
      height: 100,
      rotation: 90,
      width: 100,
    });

    expect(result.rotation).to.equal(5400000);
  });

  it("should pass flip through unchanged", () => {
    const result = createTransformation({
      flip: { horizontal: true, vertical: false },
      height: 100,
      width: 100,
    });

    expect(result.flip).to.deep.equal({ horizontal: true, vertical: false });
  });
});
