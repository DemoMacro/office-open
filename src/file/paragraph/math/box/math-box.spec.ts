import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathBox } from "./math-box";

describe("MathBox", () => {
    describe("#constructor()", () => {
        it("should create a MathBox with correct root key", () => {
            const box = new MathBox({ children: [new MathRun("text")] });
            const tree = new Formatter().format(box);

            expect(tree).to.deep.equal({
                "m:box": [
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

        it("should create a MathBox with properties", () => {
            const box = new MathBox({
                properties: { opEmu: true, diff: false },
                children: [new MathRun("x")],
            });
            const tree = new Formatter().format(box);

            expect(tree).to.deep.equal({
                "m:box": [
                    {
                        "m:boxPr": [{ "m:opEmu": {} }, { "m:diff": { _attr: { "w:val": false } } }],
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
});
