import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathEqArr } from "./math-eq-arr";

describe("MathEqArr", () => {
    describe("#constructor()", () => {
        it("should create a MathEqArr with multiple rows", () => {
            const eqArr = new MathEqArr({
                rows: [[new MathRun("x + y")], [new MathRun("z = w")]],
            });
            const tree = new Formatter().format(eqArr);

            expect(tree).to.deep.equal({
                "m:eqArr": [
                    {
                        "m:e": [
                            {
                                "m:r": [{ "m:t": ["x + y"] }],
                            },
                        ],
                    },
                    {
                        "m:e": [
                            {
                                "m:r": [{ "m:t": ["z = w"] }],
                            },
                        ],
                    },
                ],
            });
        });

        it("should create a MathEqArr with properties", () => {
            const eqArr = new MathEqArr({
                properties: { maxDist: true, objDist: false },
                rows: [[new MathRun("a")]],
            });
            const tree = new Formatter().format(eqArr);

            expect(tree["m:eqArr"][0]).to.deep.equal({
                "m:eqArrPr": [{ "m:maxDist": {} }, { "m:objDist": { _attr: { "m:val": false } } }],
            });
        });
    });
});
