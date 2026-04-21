import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createBlipEffects } from "./blip-effects";

describe("BlipEffects", () => {
    describe("#createBlipEffects()", () => {
        it("should create empty array when no effects specified", () => {
            const effects = createBlipEffects({});
            expect(effects).to.deep.equal([]);
        });

        it("should create grayscale effect", () => {
            const effects = createBlipEffects({ grayscale: true });
            expect(effects.length).toBe(1);
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:grayscl": {},
            });
        });

        it("should create luminance effect with bright and contrast", () => {
            const effects = createBlipEffects({
                luminance: { bright: 30, contrast: -20 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:lum": {
                    _attr: {
                        bright: "30%",
                        contrast: "-20%",
                    },
                },
            });
        });

        it("should create HSL effect", () => {
            const effects = createBlipEffects({
                hsl: { hue: 1800000, sat: 50, lum: -10 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:hsl": {
                    _attr: {
                        hue: "1800000",
                        lum: "-10%",
                        sat: "50%",
                    },
                },
            });
        });

        it("should create tint effect", () => {
            const effects = createBlipEffects({
                tint: { hue: 6000000, amt: 40 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:tint": {
                    _attr: {
                        amt: "40%",
                        hue: "6000000",
                    },
                },
            });
        });

        it("should create duotone effect with two colors", () => {
            const effects = createBlipEffects({
                duotone: {
                    color1: { value: "000000" },
                    color2: { value: "FFFFFF" },
                },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:duotone": [
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                    {
                        "a:srgbClr": {
                            _attr: { val: "FFFFFF" },
                        },
                    },
                ],
            });
        });

        it("should create biLevel effect", () => {
            const effects = createBlipEffects({
                biLevel: { thresh: 50 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:biLevel": {
                    _attr: {
                        thresh: "50%",
                    },
                },
            });
        });

        it("should create alphaCeiling and alphaFloor effects", () => {
            const effects = createBlipEffects({
                alphaCeiling: true,
                alphaFloor: true,
            });
            expect(effects.length).toBe(2);
        });

        it("should create alphaModFix effect", () => {
            const effects = createBlipEffects({
                alphaModFix: { amount: 50 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:alphaModFix": {
                    _attr: {
                        amt: "50%",
                    },
                },
            });
        });

        it("should create alphaRepl effect", () => {
            const effects = createBlipEffects({
                alphaRepl: { amount: 75 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:alphaRepl": {
                    _attr: {
                        a: "75%",
                    },
                },
            });
        });

        it("should create alphaBiLevel effect", () => {
            const effects = createBlipEffects({
                alphaBiLevel: { thresh: 25 },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:alphaBiLevel": {
                    _attr: {
                        thresh: "25%",
                    },
                },
            });
        });

        it("should create colorChange effect", () => {
            const effects = createBlipEffects({
                colorChange: {
                    from: { value: "FF0000" },
                    to: { value: "0000FF" },
                    useAlpha: false,
                },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:clrChange": [
                    {
                        _attr: {
                            useA: "0",
                        },
                    },
                    {
                        "a:clrFrom": [
                            {
                                "a:srgbClr": {
                                    _attr: { val: "FF0000" },
                                },
                            },
                        ],
                    },
                    {
                        "a:clrTo": [
                            {
                                "a:srgbClr": {
                                    _attr: { val: "0000FF" },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create colorRepl effect", () => {
            const effects = createBlipEffects({
                colorRepl: { color: { value: "00FF00" } },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:clrRepl": [
                    {
                        "a:srgbClr": {
                            _attr: { val: "00FF00" },
                        },
                    },
                ],
            });
        });

        it("should create blur effect", () => {
            const effects = createBlipEffects({
                blur: { rad: 40000, grow: false },
            });
            const tree = new Formatter().format(effects[0]);
            expect(tree).to.deep.equal({
                "a:blur": {
                    _attr: {
                        grow: 0,
                        rad: 40000,
                    },
                },
            });
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
