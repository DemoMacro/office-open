import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathMatrix } from "./math-matrix";

describe("MathMatrix", () => {
    describe("#constructor()", () => {
        it("should create a MathMatrix with rows and columns", () => {
            const matrix = new MathMatrix({
                rows: [
                    [new MathRun("a"), new MathRun("b")],
                    [new MathRun("c"), new MathRun("d")],
                ],
            });
            const tree = new Formatter().format(matrix);

            expect(tree).to.deep.equal({
                "m:m": [
                    {
                        "m:mr": [
                            {
                                "m:e": [
                                    {
                                        "m:r": [{ "m:t": ["a"] }],
                                    },
                                ],
                            },
                            {
                                "m:e": [
                                    {
                                        "m:r": [{ "m:t": ["b"] }],
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "m:mr": [
                            {
                                "m:e": [
                                    {
                                        "m:r": [{ "m:t": ["c"] }],
                                    },
                                ],
                            },
                            {
                                "m:e": [
                                    {
                                        "m:r": [{ "m:t": ["d"] }],
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it("should create a MathMatrix with properties", () => {
            const matrix = new MathMatrix({
                properties: { plcHide: true },
                rows: [[new MathRun("x")]],
            });
            const tree = new Formatter().format(matrix);

            expect(tree["m:m"][0]).to.deep.equal({
                "m:mPr": [{ "m:plcHide": {} }],
            });
        });
    });
});
