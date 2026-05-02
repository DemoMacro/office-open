import { describe, expect, it } from "vite-plus/test";

import { convertInchesToTwip, convertMillimetersToTwip } from "./converters";

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
