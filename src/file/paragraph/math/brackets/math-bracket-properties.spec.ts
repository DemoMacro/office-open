import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathBracketProperties } from "./math-bracket-properties";

describe("createMathBracketProperties", () => {
    describe("#constructor()", () => {
        it("should create a MathBracketProperties with correct root key", () => {
            const mathBracketProperties = createMathBracketProperties({});

            const tree = new Formatter().format(mathBracketProperties);
            expect(tree).to.deep.equal({
                "m:dPr": {},
            });
        });

        it("should create a MathBracketProperties with correct root key and add brackets", () => {
            const mathBracketProperties = createMathBracketProperties({
                characters: {
                    beginningCharacter: "[",
                    endingCharacter: "]",
                },
            });

            const tree = new Formatter().format(mathBracketProperties);
            expect(tree).to.deep.equal({
                "m:dPr": [
                    {
                        "m:begChr": {
                            _attr: {
                                "m:val": "[",
                            },
                        },
                    },
                    {
                        "m:endChr": {
                            _attr: {
                                "m:val": "]",
                            },
                        },
                    },
                ],
            });
        });
    });
});
