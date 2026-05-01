import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathParagraph } from "./math-para";
import { MathRun } from "./math-run";

describe("MathParagraph", () => {
    it("should create a math paragraph with children", () => {
        const tree = new Formatter().format(
            new MathParagraph({
                children: [{ children: [new MathRun("x + y")] }],
            }),
        );
        expect(tree).to.deep.equal({
            "m:oMathPara": [
                {
                    "m:oMath": [
                        {
                            "m:r": [{ "m:t": ["x + y"] }],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a math paragraph with justification", () => {
        const tree = new Formatter().format(
            new MathParagraph({
                justification: "center",
                children: [{ children: [new MathRun("x")] }],
            }),
        );
        expect(tree).to.deep.equal({
            "m:oMathPara": [
                {
                    "m:oMathParaPr": [{ "m:jc": { _attr: { "m:val": "center" } } }],
                },
                {
                    "m:oMath": [
                        {
                            "m:r": [{ "m:t": ["x"] }],
                        },
                    ],
                },
            ],
        });
    });
});
