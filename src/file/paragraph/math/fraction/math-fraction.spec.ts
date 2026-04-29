import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathFraction } from "./math-fraction";

describe("MathFraction", () => {
    describe("#constructor()", () => {
        it("should create a MathFraction with correct root key", () => {
            const mathFraction = new MathFraction({
                denominator: [new MathRun("2")],
                numerator: [new MathRun("2")],
            });
            const tree = new Formatter().format(mathFraction);
            expect(tree).to.deep.equal({
                "m:f": [
                    {
                        "m:num": [
                            {
                                "m:r": [
                                    {
                                        "m:t": ["2"],
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "m:den": [
                            {
                                "m:r": [
                                    {
                                        "m:t": ["2"],
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it("should create a MathFraction with fractionType BAR", () => {
            const mathFraction = new MathFraction({
                denominator: [new MathRun("2")],
                numerator: [new MathRun("1")],
                fractionType: "BAR",
            });
            const tree = new Formatter().format(mathFraction);
            expect(tree).to.deep.equal({
                "m:f": [
                    {
                        "m:fPr": [{ "m:type": { _attr: { val: "bar" } } }],
                    },
                    {
                        "m:num": [{ "m:r": [{ "m:t": ["1"] }] }],
                    },
                    {
                        "m:den": [{ "m:r": [{ "m:t": ["2"] }] }],
                    },
                ],
            });
        });

        it("should create a MathFraction with fractionType SKEWED", () => {
            const mathFraction = new MathFraction({
                denominator: [new MathRun("b")],
                numerator: [new MathRun("a")],
                fractionType: "SKEWED",
            });
            const tree = new Formatter().format(mathFraction);
            expect(tree["m:f"][0]).to.deep.equal({
                "m:fPr": [{ "m:type": { _attr: { val: "skw" } } }],
            });
        });

        it("should create a MathFraction with fractionType LINEAR", () => {
            const mathFraction = new MathFraction({
                denominator: [new MathRun("b")],
                numerator: [new MathRun("a")],
                fractionType: "LINEAR",
            });
            const tree = new Formatter().format(mathFraction);
            expect(tree["m:f"][0]).to.deep.equal({
                "m:fPr": [{ "m:type": { _attr: { val: "lin" } } }],
            });
        });

        it("should create a MathFraction with fractionType NO_BAR", () => {
            const mathFraction = new MathFraction({
                denominator: [new MathRun("b")],
                numerator: [new MathRun("a")],
                fractionType: "NO_BAR",
            });
            const tree = new Formatter().format(mathFraction);
            expect(tree["m:f"][0]).to.deep.equal({
                "m:fPr": [{ "m:type": { _attr: { val: "noBar" } } }],
            });
        });
    });
});
