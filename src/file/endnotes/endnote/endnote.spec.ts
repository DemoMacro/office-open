import { Formatter } from "@export/formatter";
import { Paragraph, TextRun } from "@file/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { Endnote, EndnoteType } from "./endnote";

describe("Endnote", () => {
    describe("#constructor", () => {
        it("should create an endnote with an endnote type", () => {
            const endnote = new Endnote({
                children: [],
                id: 1,
                type: EndnoteType.SEPARATOR,
            });
            const tree = new Formatter().format(endnote);

            expect(Object.keys(tree)).to.deep.equal(["w:endnote"]);
            expect(tree["w:endnote"]).to.deep.equal({
                _attr: { "w:id": 1, "w:type": "separator" },
            });
        });

        it("should create a endnote without a endnote type", () => {
            const endnote = new Endnote({
                children: [],
                id: 1,
            });
            const tree = new Formatter().format(endnote);

            expect(Object.keys(tree)).to.deep.equal(["w:endnote"]);
            expect(tree["w:endnote"]).to.deep.equal({ _attr: { "w:id": 1 } });
        });

        it("should append endnote ref run on the first endnote paragraph", () => {
            const endnote = new Endnote({
                children: [new Paragraph({ children: [new TextRun("test-endnote")] })],
                id: 1,
            });
            const tree = new Formatter().format(endnote);

            expect(tree).to.deep.equal({
                "w:endnote": [
                    {
                        _attr: {
                            "w:id": 1,
                        },
                    },
                    {
                        "w:p": [
                            {
                                "w:r": [
                                    {
                                        "w:rPr": [
                                            {
                                                "w:rStyle": {
                                                    _attr: {
                                                        "w:val": "EndnoteReference",
                                                    },
                                                },
                                            },
                                        ],
                                    },
                                    {
                                        "w:endnoteRef": {},
                                    },
                                ],
                            },
                            {
                                "w:r": [
                                    {
                                        "w:t": [
                                            {
                                                _attr: {
                                                    "xml:space": "preserve",
                                                },
                                            },
                                            "test-endnote",
                                        ],
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it("should add multiple paragraphs", () => {
            const endnote = new Endnote({
                children: [
                    new Paragraph({ children: [new TextRun("test-endnote")] }),
                    new Paragraph({ children: [new TextRun("test-endnote-2")] }),
                ],
                id: 1,
            });
            const tree = new Formatter().format(endnote);

            expect(tree).to.deep.equal({
                "w:endnote": [
                    {
                        _attr: {
                            "w:id": 1,
                        },
                    },
                    {
                        "w:p": [
                            {
                                "w:r": [
                                    {
                                        "w:rPr": [
                                            {
                                                "w:rStyle": {
                                                    _attr: {
                                                        "w:val": "EndnoteReference",
                                                    },
                                                },
                                            },
                                        ],
                                    },
                                    {
                                        "w:endnoteRef": {},
                                    },
                                ],
                            },
                            {
                                "w:r": [
                                    {
                                        "w:t": [
                                            {
                                                _attr: {
                                                    "xml:space": "preserve",
                                                },
                                            },
                                            "test-endnote",
                                        ],
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "w:p": [
                            {
                                "w:r": [
                                    {
                                        "w:t": [
                                            {
                                                _attr: {
                                                    "xml:space": "preserve",
                                                },
                                            },
                                            "test-endnote-2",
                                        ],
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });
    });
});
