import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { createMathAccent } from "./math-accent";

describe("createMathAccent", () => {
    it("should create a math accent without accentCharacter", () => {
        const tree = new Formatter().format(
            createMathAccent({
                children: [new MathRun("x")],
            }),
        );
        expect(tree).to.deep.equal({
            "m:acc": [
                {
                    "m:e": [
                        {
                            "m:r": [{ "m:t": ["x"] }],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a math accent with accentCharacter", () => {
        const tree = new Formatter().format(
            createMathAccent({
                accentCharacter: "\u0302",
                children: [new MathRun("x")],
            }),
        );
        expect(tree).to.deep.equal({
            "m:acc": [
                {
                    "m:accPr": [{ "m:chr": { _attr: { "m:val": "\u0302" } } }],
                },
                {
                    "m:e": [
                        {
                            "m:r": [{ "m:t": ["x"] }],
                        },
                    ],
                },
            ],
        });
    });
});
