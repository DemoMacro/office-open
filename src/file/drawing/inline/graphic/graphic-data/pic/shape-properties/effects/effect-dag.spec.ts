import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createEffectDag } from "./effect-dag";
import { BlendMode } from "./fill-overlay";

describe("createEffectDag", () => {
    it("should create an empty effectDag", () => {
        const tree = new Formatter().format(createEffectDag({}));
        expect(tree).to.deep.equal({
            "a:effectDag": {},
        });
    });

    it("should create an effectDag with type and name attributes", () => {
        const tree = new Formatter().format(createEffectDag({ type: "tree", name: "myEffects" }));
        expect(tree).to.deep.equal({
            "a:effectDag": { _attr: { type: "tree", name: "myEffects" } },
        });
    });

    // ── Existing effects (also in CT_EffectList) ──

    it("should create an effectDag with glow", () => {
        const tree = new Formatter().format(
            createEffectDag({
                glow: { rad: 50800, color: { value: "FF0000" } },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:glow": [
                        { _attr: { rad: 50800 } },
                        { "a:srgbClr": { _attr: { val: "FF0000" } } },
                    ],
                },
            ],
        });
    });

    it("should create an effectDag with outerShadow", () => {
        const tree = new Formatter().format(
            createEffectDag({
                outerShadow: {
                    blurRad: 76200,
                    dist: 38100,
                    dir: 5400000,
                    color: { value: "000000" },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:outerShdw": [
                        { _attr: { blurRad: 76200, dir: 5400000, dist: 38100 } },
                        { "a:srgbClr": { _attr: { val: "000000" } } },
                    ],
                },
            ],
        });
    });

    it("should create an effectDag with blur, reflection, and softEdge", () => {
        const tree = new Formatter().format(
            createEffectDag({
                blur: { rad: 40000 },
                reflection: true,
                softEdge: 25400,
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                { "a:blur": { _attr: { rad: 40000 } } },
                { "a:reflection": {} },
                { "a:softEdge": { _attr: { rad: 25400 } } },
            ],
        });
    });

    it("should create an effectDag with fillOverlay", () => {
        const tree = new Formatter().format(
            createEffectDag({
                fillOverlay: { blend: BlendMode.MULTIPLY, solidFill: { value: "FF0000" } },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:fillOverlay": [
                        { _attr: { blend: "mult" } },
                        {
                            "a:solidFill": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }],
                        },
                    ],
                },
            ],
        });
    });

    // ── Alpha effects ──

    it("should create an effectDag with alphaBiLevel", () => {
        const tree = new Formatter().format(createEffectDag({ alphaBiLevel: { thresh: 50000 } }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaBiLevel": { _attr: { thresh: 50000 } } }],
        });
    });

    it("should create an effectDag with alphaCeiling and alphaFloor", () => {
        const tree = new Formatter().format(
            createEffectDag({ alphaCeiling: true, alphaFloor: true }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaCeiling": {} }, { "a:alphaFloor": {} }],
        });
    });

    it("should create an effectDag with alphaInverse", () => {
        const tree = new Formatter().format(
            createEffectDag({ alphaInverse: { color: { value: "FF0000" } } }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:alphaInv": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }],
                },
            ],
        });
    });

    it("should create an effectDag with alphaInverse without color", () => {
        const tree = new Formatter().format(createEffectDag({ alphaInverse: {} }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaInv": {} }],
        });
    });

    it("should create an effectDag with alphaModulateFixed", () => {
        const tree = new Formatter().format(
            createEffectDag({ alphaModulateFixed: { amt: 75000 } }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaModFix": { _attr: { amt: 75000 } } }],
        });
    });

    it("should create an effectDag with alphaModulateFixed with no attributes", () => {
        const tree = new Formatter().format(createEffectDag({ alphaModulateFixed: {} }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaModFix": {} }],
        });
    });

    it("should create an effectDag with alphaOutset", () => {
        const tree = new Formatter().format(createEffectDag({ alphaOutset: { rad: 50800 } }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaOutset": { _attr: { rad: 50800 } } }],
        });
    });

    it("should create an effectDag with alphaReplace", () => {
        const tree = new Formatter().format(createEffectDag({ alphaReplace: { a: 50000 } }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:alphaRepl": { _attr: { a: 50000 } } }],
        });
    });

    // ── New effects ──

    it("should create an effectDag with biLevel", () => {
        const tree = new Formatter().format(createEffectDag({ biLevel: { thresh: 50000 } }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:biLevel": { _attr: { thresh: 50000 } } }],
        });
    });

    it("should create an effectDag with grayscale", () => {
        const tree = new Formatter().format(createEffectDag({ grayscale: true }));
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:grayscl": {} }],
        });
    });

    it("should create an effectDag with hsl", () => {
        const tree = new Formatter().format(
            createEffectDag({ hsl: { hue: 5400000, sat: 100000, lum: 50000 } }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:hsl": { _attr: { hue: 5400000, lum: 50000, sat: 100000 } } }],
        });
    });

    it("should create an effectDag with luminance", () => {
        const tree = new Formatter().format(
            createEffectDag({ luminance: { bright: 10000, contrast: -5000 } }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:lum": { _attr: { bright: 10000, contrast: -5000 } } }],
        });
    });

    it("should create an effectDag with tint", () => {
        const tree = new Formatter().format(
            createEffectDag({ tint: { hue: 2700000, amt: 30000 } }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:tint": { _attr: { amt: 30000, hue: 2700000 } } }],
        });
    });

    it("should create an effectDag with relativeOffset", () => {
        const tree = new Formatter().format(
            createEffectDag({ relativeOffset: { tx: 10000, ty: -5000 } }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:relOff": { _attr: { tx: 10000, ty: -5000 } } }],
        });
    });

    it("should create an effectDag with transform", () => {
        const tree = new Formatter().format(
            createEffectDag({
                transform: { sx: 120000, sy: 80000, kx: 5400000, ky: 2700000, tx: 100, ty: 200 },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:xfrm": {
                        _attr: {
                            kx: 5400000,
                            ky: 2700000,
                            sx: 120000,
                            sy: 80000,
                            tx: 100,
                            ty: 200,
                        },
                    },
                },
            ],
        });
    });

    // ── Effects with child elements ──

    it("should create an effectDag with colorReplace", () => {
        const tree = new Formatter().format(createEffectDag({ colorReplace: { value: "FF0000" } }));
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:clrRepl": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }],
                },
            ],
        });
    });

    it("should create an effectDag with duotone", () => {
        const tree = new Formatter().format(
            createEffectDag({
                duotone: {
                    color1: { value: "FF0000" },
                    color2: { value: "0000FF" },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:duotone": [
                        { "a:srgbClr": { _attr: { val: "FF0000" } } },
                        { "a:srgbClr": { _attr: { val: "0000FF" } } },
                    ],
                },
            ],
        });
    });

    it("should create an effectDag with fill", () => {
        const tree = new Formatter().format(
            createEffectDag({
                fill: { solidFill: { value: "00FF00" } },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:fill": [
                        {
                            "a:solidFill": [{ "a:srgbClr": { _attr: { val: "00FF00" } } }],
                        },
                    ],
                },
            ],
        });
    });

    // ── Container and reference effects ──

    it("should create an effectDag with nested containers", () => {
        const tree = new Formatter().format(
            createEffectDag({
                type: "tree",
                containers: [
                    {
                        glow: { color: { value: "FF0000" } },
                        outerShadow: { color: { value: "000000" } },
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                { _attr: { type: "tree" } },
                {
                    "a:cont": [
                        {
                            "a:glow": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }],
                        },
                        {
                            "a:outerShdw": [
                                { _attr: {} },
                                { "a:srgbClr": { _attr: { val: "000000" } } },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an effectDag with effect reference", () => {
        const tree = new Formatter().format(
            createEffectDag({
                effectRefs: [{ ref: "myEffect" }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [{ "a:effect": { _attr: { ref: "myEffect" } } }],
        });
    });

    it("should create an effectDag with alphaModulate (nested container)", () => {
        const tree = new Formatter().format(
            createEffectDag({
                alphaModulate: {
                    glow: { color: { value: "FF0000" } },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:alphaMod": [
                        {
                            "a:cont": [
                                {
                                    "a:glow": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an effectDag with blend (nested container)", () => {
        const tree = new Formatter().format(
            createEffectDag({
                blend: {
                    blend: BlendMode.SCREEN,
                    container: {
                        blur: { rad: 40000 },
                    },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                {
                    "a:blend": [
                        { _attr: { blend: "screen" } },
                        {
                            "a:cont": [{ "a:blur": { _attr: { rad: 40000 } } }],
                        },
                    ],
                },
            ],
        });
    });

    // ── Complex combinations ──

    it("should create an effectDag with multiple effects in XSD order", () => {
        const tree = new Formatter().format(
            createEffectDag({
                type: "sib",
                blur: { rad: 40000 },
                glow: { rad: 50800, color: { value: "FF0000" } },
                outerShadow: { color: { value: "000000" } },
                alphaBiLevel: { thresh: 50000 },
                grayscale: true,
                luminance: { bright: 10000 },
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                { _attr: { type: "sib" } },
                { "a:blur": { _attr: { rad: 40000 } } },
                {
                    "a:glow": [
                        { _attr: { rad: 50800 } },
                        { "a:srgbClr": { _attr: { val: "FF0000" } } },
                    ],
                },
                {
                    "a:outerShdw": [{ _attr: {} }, { "a:srgbClr": { _attr: { val: "000000" } } }],
                },
                { "a:alphaBiLevel": { _attr: { thresh: 50000 } } },
                { "a:grayscl": {} },
                { "a:lum": { _attr: { bright: 10000 } } },
            ],
        });
    });

    it("should create a deeply nested effectDag", () => {
        const tree = new Formatter().format(
            createEffectDag({
                type: "tree",
                containers: [
                    {
                        type: "sib",
                        glow: { color: { value: "FF0000" } },
                        containers: [
                            {
                                alphaBiLevel: { thresh: 50000 },
                            },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:effectDag": [
                { _attr: { type: "tree" } },
                {
                    "a:cont": [
                        { _attr: { type: "sib" } },
                        {
                            "a:glow": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }],
                        },
                        {
                            "a:cont": [{ "a:alphaBiLevel": { _attr: { thresh: 50000 } } }],
                        },
                    ],
                },
            ],
        });
    });
});
