import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createEffectList } from "./effect-list";
import { BlendMode } from "./fill-overlay";
import { createGlowEffect } from "./glow";
import { createInnerShadowEffect } from "./inner-shdw";
import { createOuterShadowEffect } from "./outer-shdw";
import { createPresetShadowEffect } from "./prst-shdw";
import { createReflectionEffect } from "./reflection";
import { createSoftEdgeEffect } from "./soft-edge";

describe("Effects", () => {
    describe("createGlowEffect", () => {
        it("should create glow with color and radius", () => {
            const tree = new Formatter().format(
                createGlowEffect({ rad: 50800, color: { value: "FF0000" } }),
            );
            expect(tree).to.deep.equal({
                "a:glow": [
                    { _attr: { rad: 50800 } },
                    {
                        "a:srgbClr": {
                            _attr: { val: "FF0000" },
                        },
                    },
                ],
            });
        });

        it("should create glow with only color", () => {
            const tree = new Formatter().format(createGlowEffect({ color: { value: "00FF00" } }));
            expect(tree).to.deep.equal({
                "a:glow": [
                    {
                        "a:srgbClr": {
                            _attr: { val: "00FF00" },
                        },
                    },
                ],
            });
        });
    });

    describe("createOuterShadowEffect", () => {
        it("should create outer shadow with required color", () => {
            const tree = new Formatter().format(
                createOuterShadowEffect({
                    blurRad: 76200,
                    dist: 38100,
                    dir: 5400000,
                    color: { value: "000000" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:outerShdw": [
                    {
                        _attr: {
                            blurRad: 76200,
                            dir: 5400000,
                            dist: 38100,
                        },
                    },
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });

        it("should create outer shadow with rotWithShape false", () => {
            const tree = new Formatter().format(
                createOuterShadowEffect({
                    color: { value: "000000" },
                    rotWithShape: false,
                }),
            );
            expect(tree).to.deep.equal({
                "a:outerShdw": [
                    { _attr: { rotWithShape: 0 } },
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });

        it("should create outer shadow with all optional attributes", () => {
            const tree = new Formatter().format(
                createOuterShadowEffect({
                    blurRad: 76200,
                    dist: 38100,
                    dir: 5400000,
                    sx: 120000,
                    sy: 80000,
                    kx: 5400000,
                    ky: 2700000,
                    algn: "BOTTOM_LEFT",
                    color: { value: "000000" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:outerShdw": [
                    {
                        _attr: {
                            algn: "bl", // BOTTOM_LEFT maps to "bl"
                            blurRad: 76200,
                            dir: 5400000,
                            dist: 38100,
                            kx: 5400000,
                            ky: 2700000,
                            sx: 120000,
                            sy: 80000,
                        },
                    },
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });
    });

    describe("createInnerShadowEffect", () => {
        it("should create inner shadow with only color", () => {
            const tree = new Formatter().format(
                createInnerShadowEffect({
                    color: { value: "000000" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:innerShdw": [
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });

        it("should create inner shadow with color and attributes", () => {
            const tree = new Formatter().format(
                createInnerShadowEffect({
                    blurRad: 50800,
                    dist: 38100,
                    dir: 2700000,
                    color: { value: "000000" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:innerShdw": [
                    {
                        _attr: {
                            blurRad: 50800,
                            dir: 2700000,
                            dist: 38100,
                        },
                    },
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });
    });

    describe("createPresetShadowEffect", () => {
        it("should create preset shadow with only prst and color", () => {
            const tree = new Formatter().format(
                createPresetShadowEffect({
                    prst: "SHDW2",
                    color: { value: "000000" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:prstShdw": [
                    { _attr: { prst: "shdw2" } },
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });

        it("should create preset shadow with prst, dist, dir and color", () => {
            const tree = new Formatter().format(
                createPresetShadowEffect({
                    prst: "SHDW2",
                    dist: 38100,
                    dir: 5400000,
                    color: { value: "000000" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:prstShdw": [
                    {
                        _attr: {
                            dir: 5400000,
                            dist: 38100,
                            prst: "shdw2",
                        },
                    },
                    {
                        "a:srgbClr": {
                            _attr: { val: "000000" },
                        },
                    },
                ],
            });
        });
    });

    describe("createReflectionEffect", () => {
        it("should create reflection with no options", () => {
            const tree = new Formatter().format(createReflectionEffect());
            expect(tree).to.deep.equal({
                "a:reflection": {},
            });
        });

        it("should create reflection with attributes", () => {
            const tree = new Formatter().format(
                createReflectionEffect({
                    blurRad: 6350,
                    stA: 40000,
                    dist: 38100,
                    fadeDir: 5400000,
                }),
            );
            expect(tree).to.deep.equal({
                "a:reflection": {
                    _attr: {
                        blurRad: 6350,
                        dist: 38100,
                        fadeDir: 5400000,
                        stA: 40000,
                    },
                },
            });
        });

        it("should create reflection with all optional attributes", () => {
            const tree = new Formatter().format(
                createReflectionEffect({
                    blurRad: 6350,
                    stA: 40000,
                    stPos: 10000,
                    endA: 20000,
                    endPos: 90000,
                    dist: 38100,
                    dir: 5400000,
                    fadeDir: 5400000,
                    sx: 120000,
                    sy: 80000,
                    kx: 5400000,
                    ky: 2700000,
                    algn: "bl",
                    rotWithShape: false,
                }),
            );
            expect(tree).to.deep.equal({
                "a:reflection": {
                    _attr: {
                        algn: "bl",
                        blurRad: 6350,
                        dir: 5400000,
                        dist: 38100,
                        endA: 20000,
                        endPos: 90000,
                        fadeDir: 5400000,
                        kx: 5400000,
                        ky: 2700000,
                        rotWithShape: 0,
                        stA: 40000,
                        stPos: 10000,
                        sx: 120000,
                        sy: 80000,
                    },
                },
            });
        });
    });

    describe("createSoftEdgeEffect", () => {
        it("should create soft edge with radius", () => {
            const tree = new Formatter().format(createSoftEdgeEffect(50800));
            expect(tree).to.deep.equal({
                "a:softEdge": {
                    _attr: {
                        rad: 50800,
                    },
                },
            });
        });
    });

    describe("createEffectList", () => {
        it("should create effect list with blur effect with no attributes", () => {
            const tree = new Formatter().format(
                createEffectList({
                    blur: {},
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [{ "a:blur": {} }],
            });
        });

        it("should create effect list with blur effect", () => {
            const tree = new Formatter().format(
                createEffectList({
                    blur: { rad: 40000, grow: false },
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    {
                        "a:blur": {
                            _attr: {
                                grow: 0,
                                rad: 40000,
                            },
                        },
                    },
                ],
            });
        });

        it("should create effect list with inner shadow", () => {
            const tree = new Formatter().format(
                createEffectList({
                    innerShdw: {
                        blurRad: 50800,
                        dist: 38100,
                        dir: 2700000,
                        color: { value: "000000" },
                    },
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    {
                        "a:innerShdw": [
                            {
                                _attr: {
                                    blurRad: 50800,
                                    dir: 2700000,
                                    dist: 38100,
                                },
                            },
                            {
                                "a:srgbClr": {
                                    _attr: { val: "000000" },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create effect list with preset shadow", () => {
            const tree = new Formatter().format(
                createEffectList({
                    prstShdw: {
                        prst: "SHDW5",
                        dist: 38100,
                        dir: 5400000,
                        color: { value: "000000" },
                    },
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    {
                        "a:prstShdw": [
                            {
                                _attr: {
                                    dir: 5400000,
                                    dist: 38100,
                                    prst: "shdw5",
                                },
                            },
                            {
                                "a:srgbClr": {
                                    _attr: { val: "000000" },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create effect list with single effect", () => {
            const tree = new Formatter().format(
                createEffectList({
                    glow: { rad: 50800, color: { value: "FF0000" } },
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    {
                        "a:glow": [
                            { _attr: { rad: 50800 } },
                            {
                                "a:srgbClr": {
                                    _attr: { val: "FF0000" },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create effect list with multiple effects in XSD order", () => {
            const tree = new Formatter().format(
                createEffectList({
                    outerShdw: {
                        blurRad: 76200,
                        dist: 38100,
                        dir: 5400000,
                        color: { value: "000000" },
                    },
                    reflection: true,
                    softEdge: 25400,
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    {
                        "a:outerShdw": [
                            {
                                _attr: {
                                    blurRad: 76200,
                                    dir: 5400000,
                                    dist: 38100,
                                },
                            },
                            {
                                "a:srgbClr": {
                                    _attr: { val: "000000" },
                                },
                            },
                        ],
                    },
                    { "a:reflection": {} },
                    {
                        "a:softEdge": {
                            _attr: { rad: 25400 },
                        },
                    },
                ],
            });
        });

        it("should create effect list with reflection options", () => {
            const tree = new Formatter().format(
                createEffectList({
                    reflection: { blurRad: 6350, stA: 40000 },
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    {
                        "a:reflection": {
                            _attr: {
                                blurRad: 6350,
                                stA: 40000,
                            },
                        },
                    },
                ],
            });
        });

        it("should create effect list with fillOverlay between blur and glow", () => {
            const tree = new Formatter().format(
                createEffectList({
                    blur: {},
                    fillOverlay: { blend: BlendMode.MULTIPLY, color: { value: "FF0000" } },
                    glow: { color: { value: "00FF00" } },
                }),
            );
            expect(tree).to.deep.equal({
                "a:effectLst": [
                    { "a:blur": {} },
                    {
                        "a:fillOverlay": [
                            { _attr: { blend: "mult" } },
                            {
                                "a:solidFill": [
                                    {
                                        "a:srgbClr": {
                                            _attr: { val: "FF0000" },
                                        },
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "a:glow": [
                            {
                                "a:srgbClr": {
                                    _attr: { val: "00FF00" },
                                },
                            },
                        ],
                    },
                ],
            });
        });
    });
});
