import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import {
    TextBodyWrappingType,
    TextHorzOverflowType,
    TextVertOverflowType,
    TextVerticalType,
    VerticalAnchor,
    createBodyProperties,
} from "./body-properties";

describe("BodyProperties", () => {
    describe("#constructor()", () => {
        it("should create with default options", () => {
            const tree = new Formatter().format(createBodyProperties());

            expect(tree).to.deep.equal({
                "wps:bodyPr": {},
            });
        });

        it("should create with margins", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    margins: {
                        bottom: 200,
                        left: 300,
                        right: 400,
                        top: 100,
                    },
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": {
                    _attr: {
                        bIns: 200,
                        lIns: 300,
                        rIns: 400,
                        tIns: 100,
                    },
                },
            });
        });

        it("should create with vertical anchor", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    verticalAnchor: VerticalAnchor.CENTER,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": {
                    _attr: {
                        anchor: "ctr",
                    },
                },
            });
        });

        it("should create with noAutoFit", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    noAutoFit: true,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": [
                    {
                        "a:noAutofit": {},
                    },
                ],
            });
        });

        it("should create with all options", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    margins: {
                        bottom: 20,
                        left: 30,
                        right: 40,
                        top: 10,
                    },
                    noAutoFit: true,
                    verticalAnchor: VerticalAnchor.BOTTOM,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": [
                    {
                        _attr: {
                            anchor: "b",
                            bIns: 20,
                            lIns: 30,
                            rIns: 40,
                            tIns: 10,
                        },
                    },
                    {
                        "a:noAutofit": {},
                    },
                ],
            });
        });

        it("should create with direct anchor property", () => {
            const tree = new Formatter().format(
                createBodyProperties({ anchor: VerticalAnchor.JUSTIFY }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { anchor: "just" } },
            });
        });

        it("should create with direct margin properties", () => {
            const tree = new Formatter().format(
                createBodyProperties({ lIns: 100, tIns: 200, rIns: 300, bIns: 400 }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { bIns: 400, lIns: 100, rIns: 300, tIns: 200 } },
            });
        });

        it("should create with text vertical type", () => {
            const tree = new Formatter().format(
                createBodyProperties({ vert: TextVerticalType.VERTICAL }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { vert: "vert" } },
            });
        });

        it("should create with text wrapping type", () => {
            const tree = new Formatter().format(
                createBodyProperties({ wrap: TextBodyWrappingType.NONE }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { wrap: "none" } },
            });
        });

        it("should create with overflow settings", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    vertOverflow: TextVertOverflowType.CLIP,
                    horzOverflow: TextHorzOverflowType.CLIP,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { horzOverflow: "clip", vertOverflow: "clip" } },
            });
        });

        it("should create with column settings", () => {
            const tree = new Formatter().format(
                createBodyProperties({ numCol: 3, spcCol: 100000 }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { numCol: 3, spcCol: 100000 } },
            });
        });

        it("should create with rotation", () => {
            const tree = new Formatter().format(createBodyProperties({ rotation: 5400000 }));

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { rot: 5400000 } },
            });
        });

        it("should create with normAutofit", () => {
            const tree = new Formatter().format(
                createBodyProperties({ normAutofit: { fontScale: 80000, lnSpcReduction: 10000 } }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": [
                    { "a:normAutofit": { _attr: { fontScale: 80000, lnSpcReduction: 10000 } } },
                ],
            });
        });

        it("should create with spAutoFit", () => {
            const tree = new Formatter().format(createBodyProperties({ spAutoFit: true }));

            expect(tree).to.deep.equal({
                "wps:bodyPr": [{ "a:spAutoFit": {} }],
            });
        });

        it("should create with prstTxWarp", () => {
            const tree = new Formatter().format(
                createBodyProperties({ prstTxWarp: { preset: "textArchUp" } }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": [{ "a:prstTxWarp": { _attr: { prst: "textArchUp" } } }],
            });
        });

        it("should create with flatTx", () => {
            const tree = new Formatter().format(createBodyProperties({ flatTx: { z: 0 } }));

            expect(tree).to.deep.equal({
                "wps:bodyPr": [{ "a:flatTx": { _attr: { z: 0 } } }],
            });
        });

        it("should create with boolean attributes", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    rtlCol: true,
                    fromWordArt: true,
                    anchorCtr: true,
                    forceAA: true,
                    upright: true,
                    compatLnSpc: true,
                    spcFirstLastPara: true,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": {
                    _attr: {
                        anchorCtr: true,
                        compatLnSpc: true,
                        forceAA: true,
                        fromWordArt: true,
                        rtlCol: true,
                        spcFirstLastPara: true,
                        upright: true,
                    },
                },
            });
        });

        it("should prefer anchor over verticalAnchor", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    anchor: VerticalAnchor.TOP,
                    verticalAnchor: VerticalAnchor.BOTTOM,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { anchor: "t" } },
            });
        });

        it("should prefer direct margin properties over margins shorthand", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    lIns: 500,
                    margins: { left: 100, top: 200, right: 300, bottom: 400 },
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": { _attr: { bIns: 400, lIns: 500, rIns: 300, tIns: 200 } },
            });
        });

        it("should create with multiple child elements in XSD order", () => {
            const tree = new Formatter().format(
                createBodyProperties({
                    prstTxWarp: { preset: "textCircle" },
                    normAutofit: { fontScale: 90000 },
                    flatTx: { z: 100 },
                    anchor: VerticalAnchor.CENTER,
                    lIns: 100,
                    tIns: 100,
                    rIns: 100,
                    bIns: 100,
                }),
            );

            expect(tree).to.deep.equal({
                "wps:bodyPr": [
                    { _attr: { anchor: "ctr", bIns: 100, lIns: 100, rIns: 100, tIns: 100 } },
                    { "a:prstTxWarp": { _attr: { prst: "textCircle" } } },
                    { "a:normAutofit": { _attr: { fontScale: 90000 } } },
                    { "a:flatTx": { _attr: { z: 100 } } },
                ],
            });
        });
    });
});
