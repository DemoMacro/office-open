import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathGroupChr } from "./math-group-chr";

describe("MathGroupChr", () => {
    describe("#constructor()", () => {
        it("should create a MathGroupChr with correct root key", () => {
            const groupChr = new MathGroupChr({ children: [new MathRun("text")] });
            const tree = new Formatter().format(groupChr);

            expect(tree).to.deep.equal({
                "m:groupChr": [
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

        it("should create a MathGroupChr with properties", () => {
            const groupChr = new MathGroupChr({
                properties: { chr: "\u23DF", pos: "bot", vertJc: "bot" },
                children: [new MathRun("x")],
            });
            const tree = new Formatter().format(groupChr);

            expect(tree).to.deep.equal({
                "m:groupChr": [
                    {
                        "m:groupChrPr": [
                            { "m:chr": { _attr: { "m:val": "\u23DF" } } },
                            { "m:pos": { _attr: { "m:val": "bot" } } },
                            { "m:vertJc": { _attr: { "m:val": "bot" } } },
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
    });
});
