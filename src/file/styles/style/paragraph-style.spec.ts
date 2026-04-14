import { Formatter } from "@export/formatter";
import { AlignmentType, EmphasisMarkType, TabStopPosition } from "@file/paragraph";
import { HighlightColor } from "@file/paragraph/run";
import { UnderlineType } from "@file/paragraph/run/underline";
import { ShadingType } from "@file/shading";
import { EMPTY_OBJECT } from "@file/xml-components";
import { describe, expect, it } from "vite-plus/test";

import { StyleForParagraph } from "./paragraph-style";

describe("ParagraphStyle", () => {
    describe("#constructor", () => {
        it("should set the style type to paragraph and use the given style id", () => {
            const style = new StyleForParagraph({ id: "myStyleId" });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
            });
        });

        it("should set the name of the style, if given", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                name: "Style Name",
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    { "w:name": { _attr: { "w:val": "Style Name" } } },
                ],
            });
        });
    });

    describe("formatting methods: style attributes", () => {
        it("#basedOn", () => {
            const style = new StyleForParagraph({ basedOn: "otherId", id: "myStyleId" });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    { "w:basedOn": { _attr: { "w:val": "otherId" } } },
                ],
            });
        });

        it("#quickFormat", () => {
            const style = new StyleForParagraph({ id: "myStyleId", quickFormat: true });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    {
                        _attr: {
                            "w:styleId": "myStyleId",
                            "w:type": "paragraph",
                        },
                    },
                    { "w:qFormat": EMPTY_OBJECT },
                ],
            });
        });

        it("#next", () => {
            const style = new StyleForParagraph({ id: "myStyleId", next: "otherId" });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    { "w:next": { _attr: { "w:val": "otherId" } } },
                ],
            });
        });
    });

    describe("formatting methods: paragraph properties", () => {
        it("#indent", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    indent: { left: 720, right: 500 },
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [{ "w:ind": { _attr: { "w:left": 720, "w:right": 500 } } }],
                    },
                ],
            });
        });

        it("#spacing", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: { spacing: { after: 150, before: 50 } },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [{ "w:spacing": { _attr: { "w:after": 150, "w:before": 50 } } }],
                    },
                ],
            });
        });

        it("#center", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    alignment: AlignmentType.CENTER,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [{ "w:jc": { _attr: { "w:val": "center" } } }],
                    },
                ],
            });
        });

        it("#character spacing", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    characterSpacing: 24,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:spacing": { _attr: { "w:val": 24 } } }],
                    },
                ],
            });
        });

        it("#left", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    alignment: AlignmentType.LEFT,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [{ "w:jc": { _attr: { "w:val": "left" } } }],
                    },
                ],
            });
        });

        it("#right", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    alignment: AlignmentType.RIGHT,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [{ "w:jc": { _attr: { "w:val": "right" } } }],
                    },
                ],
            });
        });

        it("#justified", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    alignment: AlignmentType.JUSTIFIED,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [{ "w:jc": { _attr: { "w:val": "both" } } }],
                    },
                ],
            });
        });

        it("#thematicBreak", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    thematicBreak: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [
                            {
                                "w:pBdr": [
                                    {
                                        "w:bottom": {
                                            _attr: {
                                                "w:color": "auto",
                                                "w:space": 1,
                                                "w:sz": 6,
                                                "w:val": "single",
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

        it("#contextualSpacing", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    contextualSpacing: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [
                            {
                                "w:contextualSpacing": {},
                            },
                        ],
                    },
                ],
            });
        });

        it("#leftTabStop", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    leftTabStop: 1200,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [
                            {
                                "w:tabs": [
                                    { "w:tab": { _attr: { "w:pos": 1200, "w:val": "left" } } },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it("#maxRightTabStop", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    rightTabStop: TabStopPosition.MAX,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:pPr": [
                            {
                                "w:tabs": [
                                    { "w:tab": { _attr: { "w:pos": 9026, "w:val": "right" } } },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it("#keepLines", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    keepLines: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    {
                        _attr: {
                            "w:styleId": "myStyleId",
                            "w:type": "paragraph",
                        },
                    },
                    { "w:pPr": [{ "w:keepLines": EMPTY_OBJECT }] },
                ],
            });
        });

        it("#keepNext", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    keepNext: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    {
                        _attr: {
                            "w:styleId": "myStyleId",
                            "w:type": "paragraph",
                        },
                    },
                    { "w:pPr": [{ "w:keepNext": EMPTY_OBJECT }] },
                ],
            });
        });

        it("#outlineLevel", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                paragraph: {
                    outlineLevel: 1,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    { "w:pPr": [{ "w:outlineLvl": { _attr: { "w:val": 1 } } }] },
                ],
            });
        });
    });

    describe("formatting methods: run properties", () => {
        const sizeTests = [
            {
                expected: [
                    { "w:sz": { _attr: { "w:val": 24 } } },
                    { "w:szCs": { _attr: { "w:val": 24 } } },
                ],
                size: 24,
            },
            {
                expected: [
                    { "w:sz": { _attr: { "w:val": 24 } } },
                    { "w:szCs": { _attr: { "w:val": 24 } } },
                ],
                size: 24,
                sizeComplexScript: true,
            },
            {
                expected: [{ "w:sz": { _attr: { "w:val": 24 } } }],
                size: 24,
                sizeComplexScript: false,
            },
            {
                expected: [
                    { "w:sz": { _attr: { "w:val": 24 } } },
                    { "w:szCs": { _attr: { "w:val": 26 } } },
                ],
                size: 24,
                sizeComplexScript: 26,
            },
        ];
        sizeTests.forEach(({ size, sizeComplexScript, expected }) => {
            it(`#size ${size} cs ${sizeComplexScript}`, () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: { size, sizeComplexScript },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        { "w:rPr": expected },
                    ],
                });
            });
        });

        it("#smallCaps", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    smallCaps: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:smallCaps": {} }],
                    },
                ],
            });
        });

        it("#allCaps", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    allCaps: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:caps": {} }],
                    },
                ],
            });
        });

        it("#strike", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    strike: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:strike": {} }],
                    },
                ],
            });
        });

        it("#doubleStrike", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    doubleStrike: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:dstrike": {} }],
                    },
                ],
            });
        });

        it("#subScript", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    subScript: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:vertAlign": { _attr: { "w:val": "subscript" } } }],
                    },
                ],
            });
        });

        it("#superScript", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    superScript: true,
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:vertAlign": { _attr: { "w:val": "superscript" } } }],
                    },
                ],
            });
        });

        it("#font by name", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    font: "Times",
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [
                            {
                                "w:rFonts": {
                                    _attr: {
                                        "w:ascii": "Times",
                                        "w:cs": "Times",
                                        "w:eastAsia": "Times",
                                        "w:hAnsi": "Times",
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("#font for ascii and eastAsia", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    font: {
                        ascii: "Times",
                        eastAsia: "KaiTi",
                    },
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [
                            {
                                "w:rFonts": {
                                    _attr: {
                                        "w:ascii": "Times",
                                        "w:eastAsia": "KaiTi",
                                    },
                                },
                            },
                        ],
                    },
                ],
            });
        });

        const boldTests = [
            {
                bold: true,
                expected: [{ "w:b": {} }, { "w:bCs": {} }],
            },
            {
                bold: true,
                boldComplexScript: true,
                expected: [{ "w:b": {} }, { "w:bCs": {} }],
            },
            {
                bold: true,
                boldComplexScript: false,
                expected: [{ "w:b": {} }],
            },
        ];
        boldTests.forEach(({ bold, boldComplexScript, expected }) => {
            it(`#bold ${bold} cs ${boldComplexScript}`, () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: { bold, boldComplexScript },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        { "w:rPr": expected },
                    ],
                });
            });
        });

        const italicsTests = [
            {
                expected: [{ "w:i": {} }, { "w:iCs": {} }],
                italics: true,
            },
            {
                expected: [{ "w:i": {} }, { "w:iCs": {} }],
                italics: true,
                italicsComplexScript: true,
            },
            {
                expected: [{ "w:i": {} }],
                italics: true,
                italicsComplexScript: false,
            },
        ];
        italicsTests.forEach(({ italics, italicsComplexScript, expected }) => {
            it(`#italics ${italics} cs ${italicsComplexScript}`, () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: { italics, italicsComplexScript },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        { "w:rPr": expected },
                    ],
                });
            });
        });

        const highlightTests = [
            {
                expected: [
                    { "w:highlight": { _attr: { "w:val": "yellow" } } },
                    { "w:highlightCs": { _attr: { "w:val": "yellow" } } },
                ],
                highlight: HighlightColor.YELLOW,
            },
            {
                expected: [
                    { "w:highlight": { _attr: { "w:val": "yellow" } } },
                    { "w:highlightCs": { _attr: { "w:val": "yellow" } } },
                ],
                highlight: HighlightColor.YELLOW,
                highlightComplexScript: true,
            },
            {
                expected: [{ "w:highlight": { _attr: { "w:val": "yellow" } } }],
                highlight: HighlightColor.YELLOW,
                highlightComplexScript: false,
            },
            {
                expected: [
                    { "w:highlight": { _attr: { "w:val": "yellow" } } },
                    { "w:highlightCs": { _attr: { "w:val": "550099" } } },
                ],
                highlight: HighlightColor.YELLOW,
                highlightComplexScript: "550099",
            },
        ];
        highlightTests.forEach(({ highlight, highlightComplexScript, expected }) => {
            it(`#highlight ${highlight} cs ${highlightComplexScript}`, () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: { highlight, highlightComplexScript },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        { "w:rPr": expected },
                    ],
                });
            });
        });

        const shadingTests = [
            {
                expected: [
                    {
                        "w:shd": {
                            _attr: { "w:color": "FF0000", "w:fill": "00FFFF", "w:val": "pct10" },
                        },
                    },
                ],
                shading: {
                    color: "FF0000",
                    fill: "00FFFF",
                    type: ShadingType.PERCENT_10,
                },
            },
            {
                expected: [
                    {
                        "w:shd": {
                            _attr: {
                                "w:color": "0000FF",
                                "w:fill": "0066FF",
                                "w:val": "diagCross",
                            },
                        },
                    },
                ],
                shading: {
                    color: "0000FF",
                    fill: "0066FF",
                    type: ShadingType.DIAGONAL_CROSS,
                },
            },
        ];
        shadingTests.forEach(({ shading, expected }) => {
            it("#shade correctly", () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: { shading },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        { "w:rPr": expected },
                    ],
                });
            });
        });

        describe("#underline", () => {
            it("should set underline to 'single' if no arguments are given", () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: {
                        underline: {},
                    },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        {
                            "w:rPr": [{ "w:u": { _attr: { "w:val": "single" } } }],
                        },
                    ],
                });
            });

            it("should set the style if given", () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: {
                        underline: {
                            type: UnderlineType.DOUBLE,
                        },
                    },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        {
                            "w:rPr": [{ "w:u": { _attr: { "w:val": "double" } } }],
                        },
                    ],
                });
            });

            it("should set the style and color if given", () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: {
                        underline: {
                            color: "005599",
                            type: UnderlineType.DOUBLE,
                        },
                    },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        {
                            "w:rPr": [
                                { "w:u": { _attr: { "w:color": "005599", "w:val": "double" } } },
                            ],
                        },
                    ],
                });
            });
        });

        describe("#emphasisMark", () => {
            it("should set emphasisMark to 'dot' if no arguments are given", () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: {
                        emphasisMark: {},
                    },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        {
                            "w:rPr": [{ "w:em": { _attr: { "w:val": "dot" } } }],
                        },
                    ],
                });
            });

            it("should set the style if given", () => {
                const style = new StyleForParagraph({
                    id: "myStyleId",
                    run: {
                        emphasisMark: {
                            type: EmphasisMarkType.DOT,
                        },
                    },
                });
                const tree = new Formatter().format(style);
                expect(tree).to.deep.equal({
                    "w:style": [
                        { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                        {
                            "w:rPr": [{ "w:em": { _attr: { "w:val": "dot" } } }],
                        },
                    ],
                });
            });
        });

        it("#color", () => {
            const style = new StyleForParagraph({
                id: "myStyleId",
                run: {
                    color: "123456",
                },
            });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:rPr": [{ "w:color": { _attr: { "w:val": "123456" } } }],
                    },
                ],
            });
        });

        it("#link", () => {
            const style = new StyleForParagraph({ id: "myStyleId", link: "MyLink" });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    {
                        _attr: {
                            "w:styleId": "myStyleId",
                            "w:type": "paragraph",
                        },
                    },
                    { "w:link": { _attr: { "w:val": "MyLink" } } },
                ],
            });
        });

        it("#semiHidden", () => {
            const style = new StyleForParagraph({ id: "myStyleId", semiHidden: true });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    {
                        _attr: {
                            "w:styleId": "myStyleId",
                            "w:type": "paragraph",
                        },
                    },
                    { "w:semiHidden": EMPTY_OBJECT },
                ],
            });
        });

        it("#uiPriority", () => {
            const style = new StyleForParagraph({ id: "myStyleId", uiPriority: 99 });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    {
                        "w:uiPriority": {
                            _attr: {
                                "w:val": 99,
                            },
                        },
                    },
                ],
            });
        });

        it("#unhideWhenUsed", () => {
            const style = new StyleForParagraph({ id: "myStyleId", unhideWhenUsed: true });
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    {
                        _attr: {
                            "w:styleId": "myStyleId",
                            "w:type": "paragraph",
                        },
                    },
                    { "w:unhideWhenUsed": EMPTY_OBJECT },
                ],
            });
        });
    });
});
