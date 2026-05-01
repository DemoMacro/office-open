import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "../math-run";
import { MathPhant } from "./math-phant";

describe("MathPhant", () => {
    describe("#constructor()", () => {
        it("should create a MathPhant with correct root key", () => {
            const phant = new MathPhant({ children: [new MathRun("text")] });
            const tree = new Formatter().format(phant);

            expect(tree).to.deep.equal({
                "m:phant": [
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

        it("should create a MathPhant with properties", () => {
            const phant = new MathPhant({
                properties: { show: true, zeroWid: false, transp: false },
                children: [new MathRun("x")],
            });
            const tree = new Formatter().format(phant);

            expect(tree).to.deep.equal({
                "m:phant": [
                    {
                        "m:phantPr": [
                            { "m:show": {} },
                            { "m:zeroWid": { _attr: { "w:val": false } } },
                            { "m:transp": { _attr: { "w:val": false } } },
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
