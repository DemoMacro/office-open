import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { VerticalAnchor, createBodyProperties } from "./body-properties";

describe("BodyProperties", () => {
    describe("#constructor()", () => {
        it("should create with default options", () => {
            const tree = new Formatter().format(createBodyProperties());

            expect(tree).to.deep.equal({
                "wps:bodyPr": {
                    _attr: {},
                },
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
                        _attr: {},
                    },
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
    });
});
