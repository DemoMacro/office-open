import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MathRun } from "./math-run";

describe("MathRun", () => {
    describe("#constructor()", () => {
        it("should create a MathRun with correct root key", () => {
            const mathRun = new MathRun("2+2");
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:t": ["2+2"],
                    },
                ],
            });
        });

        it("should create a MathRun with options object", () => {
            const mathRun = new MathRun({ text: "x" });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:t": ["x"],
                    },
                ],
            });
        });

        it("should create a MathRun with literal property", () => {
            const mathRun = new MathRun({
                text: "abc",
                properties: { lit: true },
            });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:rPr": [
                            {
                                "m:lit": {},
                            },
                        ],
                    },
                    {
                        "m:t": ["abc"],
                    },
                ],
            });
        });

        it("should create a MathRun with normal property", () => {
            const mathRun = new MathRun({
                text: "x",
                properties: { normal: true },
            });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:rPr": [
                            {
                                "m:nor": {},
                            },
                        ],
                    },
                    {
                        "m:t": ["x"],
                    },
                ],
            });
        });

        it("should create a MathRun with script and style", () => {
            const mathRun = new MathRun({
                text: "A",
                properties: { script: "fraktur", style: "b" },
            });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:rPr": [
                            {
                                "m:scr": {
                                    _attr: {
                                        val: "fraktur",
                                    },
                                },
                            },
                            {
                                "m:sty": {
                                    _attr: {
                                        val: "b",
                                    },
                                },
                            },
                        ],
                    },
                    {
                        "m:t": ["A"],
                    },
                ],
            });
        });

        it("should create a MathRun with break alignment", () => {
            const mathRun = new MathRun({
                text: "x",
                properties: { breakAlignment: 1 },
            });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:rPr": [
                            {
                                "m:brk": {
                                    _attr: {
                                        "m:alnAt": "1",
                                    },
                                },
                            },
                        ],
                    },
                    {
                        "m:t": ["x"],
                    },
                ],
            });
        });

        it("should create a MathRun with align property", () => {
            const mathRun = new MathRun({
                text: "x",
                properties: { align: true },
            });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:rPr": [
                            {
                                "m:aln": {},
                            },
                        ],
                    },
                    {
                        "m:t": ["x"],
                    },
                ],
            });
        });

        it("should prioritize normal over script/style", () => {
            const mathRun = new MathRun({
                text: "x",
                properties: { normal: true, script: "fraktur", style: "b" },
            });
            const tree = new Formatter().format(mathRun);
            expect(tree).to.deep.equal({
                "m:r": [
                    {
                        "m:rPr": [
                            {
                                "m:nor": {},
                            },
                        ],
                    },
                    {
                        "m:t": ["x"],
                    },
                ],
            });
        });
    });
});
