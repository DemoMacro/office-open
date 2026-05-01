import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathBorderBox } from "./math-border-box";

describe("MathBorderBox", () => {
    describe("#constructor()", () => {
        it("should create a MathBorderBox with correct root key", () => {
            const borderBox = new MathBorderBox({ children: [new MathRun("text")] });
            const tree = new Formatter().format(borderBox);

            expect(tree).to.deep.equal({
                "m:borderBox": [
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

        it("should create a MathBorderBox with hidden borders", () => {
            const borderBox = new MathBorderBox({
                properties: { hideTop: true, hideBot: false },
                children: [new MathRun("x")],
            });
            const tree = new Formatter().format(borderBox);

            expect(tree).to.deep.equal({
                "m:borderBox": [
                    {
                        "m:borderBoxPr": [
                            { "m:hideTop": {} },
                            { "m:hideBot": { _attr: { "w:val": false } } },
                        ],
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

        it("should create a MathBorderBox with strike-through", () => {
            const borderBox = new MathBorderBox({
                properties: { strikeH: true, strikeV: true, strikeBLTR: false, strikeTLBR: false },
                children: [new MathRun("x")],
            });
            const tree = new Formatter().format(borderBox);

            expect(tree["m:borderBox"][0]).to.deep.equal({
                "m:borderBoxPr": [
                    { "m:strikeH": {} },
                    { "m:strikeV": {} },
                    { "m:strikeBLTR": { _attr: { "w:val": false } } },
                    { "m:strikeTLBR": { _attr: { "w:val": false } } },
                ],
            });
        });
    });
});
