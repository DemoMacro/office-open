import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathFunction } from "./math-function";

describe("MathFunction", () => {
    describe("#constructor()", () => {
        it("should create a MathFunction with correct root key", () => {
            const mathFunction = new MathFunction({
                children: [new MathRun("60")],
                name: [new MathRun("sin")],
            });

            const tree = new Formatter().format(mathFunction);
            expect(tree).to.deep.equal({
                "m:func": [
                    {
                        "m:funcPr": {},
                    },
                    {
                        "m:fName": [
                            {
                                "m:r": [
                                    {
                                        "m:t": ["sin"],
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "m:e": [
                            {
                                "m:r": [
                                    {
                                        "m:t": ["60"],
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
