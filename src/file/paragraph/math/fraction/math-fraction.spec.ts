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
    });
});
