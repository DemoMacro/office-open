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

        it("should create a MathBracketProperties with separatorCharacter", () => {
            const mathBracketProperties = createMathBracketProperties({
                characters: {
                    beginningCharacter: "(",
                    endingCharacter: ")",
                },
                separatorCharacter: "|",
            });

            const tree = new Formatter().format(mathBracketProperties);
            expect(tree).to.deep.equal({
                "m:dPr": [
                    { "m:begChr": { _attr: { "m:val": "(" } } },
                    { "m:sepChr": { _attr: { "m:val": "|" } } },
                    { "m:endChr": { _attr: { "m:val": ")" } } },
                ],
            });
        });

        it("should create a MathBracketProperties with grow", () => {
            const mathBracketProperties = createMathBracketProperties({
                grow: true,
            });

            const tree = new Formatter().format(mathBracketProperties);
            expect(tree).to.deep.equal({
                "m:dPr": [{ "m:grow": {} }],
            });
        });

        it("should create a MathBracketProperties with shape", () => {
            const mathBracketProperties = createMathBracketProperties({
                shape: "match",
            });

            const tree = new Formatter().format(mathBracketProperties);
            expect(tree).to.deep.equal({
                "m:dPr": [{ "m:shp": { _attr: { "m:val": "match" } } }],
            });
        });
    });
});
