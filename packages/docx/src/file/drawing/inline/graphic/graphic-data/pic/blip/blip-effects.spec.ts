import { describe, expect, it } from "vite-plus/test";

import { createBlipEffects } from "./blip-effects";

describe("BlipEffects", () => {
  describe("#createBlipEffects()", () => {
    it("should create empty array when no effects specified", () => {
      const effects = createBlipEffects({});
      expect(effects).to.deep.equal([]);
    });

    it("should create alphaCeiling and alphaFloor effects", () => {
      const effects = createBlipEffects({
        alphaCeiling: true,
        alphaFloor: true,
      });
      expect(effects.length).toBe(2);
    });

    it("should create multiple effects in correct order", () => {
      const effects = createBlipEffects({
        grayscale: true,
        luminance: { contrast: 10 },
      });
      expect(effects.length).toBe(2);
    });
  });
});
