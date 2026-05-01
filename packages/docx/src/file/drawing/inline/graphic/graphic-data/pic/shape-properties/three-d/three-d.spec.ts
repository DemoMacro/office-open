import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createBevel, createBottomBevel } from "./bevel";
import { createShape3D } from "./shape-3d";

describe("3D Shape Properties", () => {
    describe("createBevel", () => {
        it("should create bevel with no options", () => {
            const tree = new Formatter().format(createBevel());
            expect(tree).to.deep.equal({
                "a:bevelT": {},
            });
        });

        it("should create bevel with all options", () => {
            const tree = new Formatter().format(
                createBevel({ w: 76200, h: 152400, prst: "CIRCLE" }),
            );
            expect(tree).to.deep.equal({
                "a:bevelT": {
                    _attr: {
                        h: 152400,
                        prst: "circle",
                        w: 76200,
                    },
                },
            });
        });

        it("should create bottom bevel", () => {
            const tree = new Formatter().format(createBottomBevel({ w: 50800, prst: "ANGLE" }));
            expect(tree).to.deep.equal({
                "a:bevelB": {
                    _attr: {
                        prst: "angle",
                        w: 50800,
                    },
                },
            });
        });

        it("should create bottom bevel with no options", () => {
            const tree = new Formatter().format(createBottomBevel());
            expect(tree).to.deep.equal({
                "a:bevelB": {},
            });
        });
    });

    describe("createShape3D", () => {
        it("should create sp3d with material and depth", () => {
            const tree = new Formatter().format(
                createShape3D({
                    z: 76200,
                    extrusionH: 38100,
                    prstMaterial: "METAL",
                }),
            );
            expect(tree).to.deep.equal({
                "a:sp3d": {
                    _attr: {
                        extrusionH: 38100,
                        prstMaterial: "metal",
                        z: 76200,
                    },
                },
            });
        });

        it("should create sp3d with bevels and colors", () => {
            const tree = new Formatter().format(
                createShape3D({
                    bevelT: { w: 76200, h: 76200, prst: "CIRCLE" },
                    bevelB: { w: 50800, h: 50800, prst: "ANGLE" },
                    extrusionColor: { value: "FF0000" },
                    contourColor: { value: "0000FF" },
                    contourW: 25400,
                }),
            );
            expect(tree).to.deep.equal({
                "a:sp3d": [
                    {
                        _attr: {
                            contourW: 25400,
                        },
                    },
                    {
                        "a:bevelT": {
                            _attr: {
                                h: 76200,
                                prst: "circle",
                                w: 76200,
                            },
                        },
                    },
                    {
                        "a:bevelB": {
                            _attr: {
                                h: 50800,
                                prst: "angle",
                                w: 50800,
                            },
                        },
                    },
                    {
                        "a:extrusionClr": [
                            {
                                "a:srgbClr": {
                                    _attr: { val: "FF0000" },
                                },
                            },
                        ],
                    },
                    {
                        "a:contourClr": [
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

        it("should create sp3d with only bevels", () => {
            const tree = new Formatter().format(
                createShape3D({
                    bevelT: { prst: "CONVEX" },
                }),
            );
            expect(tree).to.deep.equal({
                "a:sp3d": [
                    {
                        "a:bevelT": {
                            _attr: {
                                prst: "convex",
                            },
                        },
                    },
                ],
            });
        });
    });
});
