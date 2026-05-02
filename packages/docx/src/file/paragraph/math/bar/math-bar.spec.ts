import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { createMathBar } from "./math-bar";

describe("MathBar", () => {
    describe("#constructor()", () => {
        it("should create a MathBar with correct root key", () => {
            const mathBar = createMathBar({ children: [new MathRun("text")], type: "top" });
            const tree = new Formatter().format(mathBar);

            expect(tree).to.deep.equal({
                "m:bar": [
                    {
                        "m:barPr": [
                            {
                                "m:pos": {
                                    _attr: {
                                        "m:val": "top",
                                    },
                                },
                            },
                        ],
                    },
                    {
                        "m:e": [
                            {
                                "m:r": [{ "m:t": ["text"] }],
                            },
                        ],
                    },
                ],
            });
        });
    });
});
