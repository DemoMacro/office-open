import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { DocumentWrapper } from "../document-wrapper";
import type { File } from "../file";
import { FontWrapper } from "../fonts/font-wrapper";
import { AlignmentType } from "./formatting";
import { ParagraphProperties, TextAlignmentType, TextboxTightWrapType } from "./properties";

describe("ParagraphProperties", () => {
    describe("#constructor()", () => {
        it("creates an initially empty property object", () => {
            const properties = new ParagraphProperties();

            expect(() => new Formatter().format(properties)).to.throw(
                "XMLComponent did not format correctly",
            );
        });

        it("should create with numbering", () => {
            const properties = new ParagraphProperties({
                numbering: {
                    instance: 0,
                    level: 0,
                    reference: "test-reference",
                },
            });
            const tree = new Formatter().format(properties, {
                file: {
                    Numbering: {
                        createConcreteNumberingInstance: (_: string, __: number) => undefined,
                    },
                } as File,
                stack: [],
                viewWrapper: new DocumentWrapper({ background: {} }),
            });

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:pStyle": {
                            _attr: {
                                "w:val": "ListParagraph",
                            },
                        },
                    },
                    {
                        "w:numPr": [
                            {
                                "w:ilvl": {
                                    _attr: {
                                        "w:val": 0,
                                    },
                                },
                            },
                            {
                                "w:numId": {
                                    _attr: {
                                        "w:val": "{test-reference-0}",
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create with numbering disabled", () => {
            const properties = new ParagraphProperties({
                numbering: false,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:numPr": [
                            {
                                "w:ilvl": {
                                    _attr: {
                                        "w:val": 0,
                                    },
                                },
                            },
                            {
                                "w:numId": {
                                    _attr: {
                                        "w:val": 0,
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create with widowControl", () => {
            const properties = new ParagraphProperties({
                widowControl: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:widowControl": {},
                    },
                ],
            });
        });

        it("should create with the bidirectional property", () => {
            const properties = new ParagraphProperties({
                bidirectional: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:bidi": {},
                    },
                ],
            });
        });

        it("should create with the contextualSpacing property", () => {
            const properties = new ParagraphProperties({
                contextualSpacing: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:contextualSpacing": {},
                    },
                ],
            });
        });

        it("should create with the suppressLineNumbers property", () => {
            const properties = new ParagraphProperties({
                suppressLineNumbers: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:suppressLineNumbers": {},
                    },
                ],
            });
        });

        it("should create with the autoSpaceEastAsianText property", () => {
            const properties = new ParagraphProperties({
                autoSpaceEastAsianText: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:autoSpaceDN": {},
                    },
                ],
            });
        });

        it("should create with the wordWrap property", () => {
            const properties = new ParagraphProperties({
                wordWrap: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:wordWrap": {
                            _attr: {
                                "w:val": 0,
                            },
                        },
                    },
                ],
            });
        });

        it("should create with the overflowPunct property", () => {
            const properties = new ParagraphProperties({
                overflowPunctuation: true,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:overflowPunct": {},
                    },
                ],
            });
        });

        it("should create with the run property", () => {
            const properties = new ParagraphProperties({
                run: {
                    size: "10pt",
                },
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:rPr": [
                            {
                                "w:sz": {
                                    _attr: {
                                        "w:val": "10pt",
                                    },
                                },
                            },
                            {
                                "w:szCs": {
                                    _attr: {
                                        "w:val": "10pt",
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create with the run property insertion", () => {
            const properties = new ParagraphProperties({
                run: {
                    insertion: { author: "Firstname Lastname", date: "123", id: 1 },
                },
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:rPr": [
                            {
                                "w:ins": {
                                    _attr: {
                                        "w:author": "Firstname Lastname",
                                        "w:date": "123",
                                        "w:id": 1,
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create with the run property deletion", () => {
            const properties = new ParagraphProperties({
                run: {
                    deletion: { author: "Firstname Lastname", date: "123", id: 1 },
                },
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:rPr": [
                            {
                                "w:del": {
                                    _attr: {
                                        "w:author": "Firstname Lastname",
                                        "w:date": "123",
                                        "w:id": 1,
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should skip numbering instance creation when viewWrapper is FontWrapper", () => {
            const properties = new ParagraphProperties({
                numbering: {
                    instance: 0,
                    level: 0,
                    reference: "test-reference",
                },
            });
            const tree = new Formatter().format(properties, {
                file: {} as File,
                stack: [],
                viewWrapper: new FontWrapper([]),
            });

            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:pStyle": {
                            _attr: {
                                "w:val": "ListParagraph",
                            },
                        },
                    },
                    {
                        "w:numPr": [
                            {
                                "w:ilvl": {
                                    _attr: {
                                        "w:val": 0,
                                    },
                                },
                            },
                            {
                                "w:numId": {
                                    _attr: {
                                        "w:val": "{test-reference-0}",
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("should create with revision", () => {
            const properties = new ParagraphProperties({
                alignment: AlignmentType.CENTER,
                revision: {
                    alignment: AlignmentType.LEFT,
                    author: "Firstname Lastname",
                    date: "123",
                    id: 1,
                },
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [
                    {
                        "w:jc": {
                            _attr: {
                                "w:val": "center",
                            },
                        },
                    },
                    {
                        "w:pPrChange": [
                            {
                                _attr: {
                                    "w:author": "Firstname Lastname",
                                    "w:date": "123",
                                    "w:id": 1,
                                },
                            },
                            {
                                "w:pPr": [
                                    {
                                        "w:jc": {
                                            _attr: {
                                                "w:val": "left",
                                            },
                                        },
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it("should create with suppressAutoHyphens", () => {
            const properties = new ParagraphProperties({ suppressAutoHyphens: true });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:suppressAutoHyphens": {} }],
            });
        });

        it("should create with adjustRightInd", () => {
            const properties = new ParagraphProperties({ adjustRightInd: true });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:adjustRightInd": {} }],
            });
        });

        it("should create with snapToGrid", () => {
            const properties = new ParagraphProperties({ snapToGrid: true });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:snapToGrid": {} }],
            });
        });

        it("should create with mirrorIndents", () => {
            const properties = new ParagraphProperties({ mirrorIndents: true });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:mirrorIndents": {} }],
            });
        });

        it("should create with kinsoku", () => {
            const properties = new ParagraphProperties({ kinsoku: false });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:kinsoku": { _attr: { "w:val": false } } }],
            });
        });

        it("should create with topLinePunct", () => {
            const properties = new ParagraphProperties({ topLinePunct: true });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:topLinePunct": {} }],
            });
        });

        it("should create with autoSpaceDE", () => {
            const properties = new ParagraphProperties({ autoSpaceDE: true });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:autoSpaceDE": {} }],
            });
        });

        it("should create with textAlignment", () => {
            const properties = new ParagraphProperties({ textAlignment: TextAlignmentType.CENTER });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:textAlignment": { _attr: { "w:val": "center" } } }],
            });
        });

        it("should create with textboxTightWrap", () => {
            const properties = new ParagraphProperties({
                textboxTightWrap: TextboxTightWrapType.ALL_LINES,
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:pPr": [{ "w:textboxTightWrap": { _attr: { "w:val": "allLines" } } }],
            });
        });
    });
});
