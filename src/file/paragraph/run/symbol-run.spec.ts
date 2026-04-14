import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { EmphasisMarkType } from "./emphasis-mark";
import { SymbolRun } from "./symbol-run";
import { UnderlineType } from "./underline";

describe("SymbolRun", () => {
    let run: SymbolRun;

    describe("#constructor()", () => {
        it("should create symbol run from text input", () => {
            run = new SymbolRun("F071");
            const f = new Formatter().format(run);
            expect(f).to.deep.equal({
                "w:r": [{ "w:sym": { _attr: { "w:char": "F071", "w:font": "Wingdings" } } }],
            });
        });

        it("should create symbol run from object input with just 'char' specified", () => {
            run = new SymbolRun({ char: "F071" });
            const f = new Formatter().format(run);
            expect(f).to.deep.equal({
                "w:r": [{ "w:sym": { _attr: { "w:char": "F071", "w:font": "Wingdings" } } }],
            });
        });

        it("should create symbol run from object input with just 'char' specified", () => {
            run = new SymbolRun({ char: "F071", symbolfont: "Arial" });
            const f = new Formatter().format(run);
            expect(f).to.deep.equal({
                "w:r": [{ "w:sym": { _attr: { "w:char": "F071", "w:font": "Arial" } } }],
            });
        });

        it("should add other standard run properties", () => {
            run = new SymbolRun({
                bold: true,
                char: "F071",
                color: "00FF00",
                emphasisMark: {
                    type: EmphasisMarkType.DOT,
                },
                highlight: "yellow",
                italics: true,
                size: 40,
                symbolfont: "Arial",
                underline: {
                    color: "ff0000",
                    type: UnderlineType.DOUBLE,
                },
            });

            const f = new Formatter().format(run);
            expect(f).to.deep.equal({
                "w:r": [
                    {
                        "w:rPr": [
                            { "w:b": {} },
                            { "w:bCs": {} },
                            { "w:i": {} },
                            { "w:iCs": {} },
                            { "w:color": { _attr: { "w:val": "00FF00" } } },
                            { "w:sz": { _attr: { "w:val": 40 } } },
                            { "w:szCs": { _attr: { "w:val": 40 } } },
                            { "w:highlight": { _attr: { "w:val": "yellow" } } },
                            { "w:highlightCs": { _attr: { "w:val": "yellow" } } },
                            { "w:u": { _attr: { "w:color": "ff0000", "w:val": "double" } } },
                            { "w:em": { _attr: { "w:val": "dot" } } },
                        ],
                    },
                    {
                        "w:sym": { _attr: { "w:char": "F071", "w:font": "Arial" } },
                    },
                ],
            });
        });
    });
});
