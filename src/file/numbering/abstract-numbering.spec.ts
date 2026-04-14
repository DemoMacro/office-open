import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { AlignmentType, EmphasisMarkType, TabStopPosition } from "../paragraph";
import { HighlightColor } from "../paragraph/run/properties";
import { UnderlineType } from "../paragraph/run/underline";
import { ShadingType } from "../shading";
import { AbstractNumbering } from "./abstract-numbering";
import { LevelFormat, LevelSuffix } from "./level";

describe("AbstractNumbering", () => {
    it("stores its ID at its .id property", () => {
        const abstractNumbering = new AbstractNumbering(5, []);
        expect(abstractNumbering.id).to.equal(5);
    });

    describe("#createLevel", () => {
        it("creates a level with the given characteristics", () => {
            const abstractNumbering = new AbstractNumbering(1, [
                {
                    alignment: AlignmentType.END,
                    format: LevelFormat.LOWER_LETTER,
                    level: 3,
                    text: "%1)",
                },
            ]);
            const tree = new Formatter().format(abstractNumbering);
            expect(tree).to.deep.equal({
                "w:abstractNum": [
                    {
                        _attr: {
                            "w15:restartNumberingAfterBreak": 0,
                            "w:abstractNumId": 1,
                        },
                    },
                    {
                        "w:multiLevelType": {
                            _attr: {
                                "w:val": "hybridMultilevel",
                            },
                        },
                    },
                    {
                        "w:lvl": [
                            {
                                "w:start": {
                                    _attr: {
                                        "w:val": 1,
                                    },
                                },
                            },
                            {
                                "w:numFmt": {
                                    _attr: {
                                        "w:val": LevelFormat.LOWER_LETTER,
                                    },
                                },
                            },
                            {
                                "w:lvlText": {
                                    _attr: {
                                        "w:val": "%1)",
                                    },
                                },
                            },
                            {
                                "w:lvlJc": {
                                    _attr: {
                                        "w:val": "end",
                                    },
                                },
                            },
                            {
                                _attr: {
                                    "w15:tentative": 1,
                                    "w:ilvl": 3,
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("uses 'start' as the default alignment", () => {
            const abstractNumbering = new AbstractNumbering(1, [
                {
                    format: LevelFormat.LOWER_LETTER,
                    level: 3,
                    text: "%1)",
                },
            ]);
            const tree = new Formatter().format(abstractNumbering);
            expect(tree).to.deep.equal({
                "w:abstractNum": [
                    {
                        _attr: {
                            "w15:restartNumberingAfterBreak": 0,
                            "w:abstractNumId": 1,
                        },
                    },
                    {
                        "w:multiLevelType": {
                            _attr: {
                                "w:val": "hybridMultilevel",
                            },
                        },
                    },
                    {
                        "w:lvl": [
                            {
                                "w:start": {
                                    _attr: {
                                        "w:val": 1,
                                    },
                                },
                            },
                            {
                                "w:numFmt": {
                                    _attr: {
                                        "w:val": LevelFormat.LOWER_LETTER,
                                    },
                                },
                            },
                            {
                                "w:lvlText": {
                                    _attr: {
                                        "w:val": "%1)",
                                    },
                                },
                            },
                            {
                                "w:lvlJc": {
                                    _attr: {
                                        "w:val": "start",
                                    },
                                },
                            },
                            {
                                _attr: {
                                    "w15:tentative": 1,
                                    "w:ilvl": 3,
                                },
                            },
                        ],
                    },
                ],
            });
        });

        it("has suffix", () => {
            const abstractNumbering = new AbstractNumbering(1, [
                {
                    alignment: AlignmentType.END,
                    format: LevelFormat.LOWER_LETTER,
                    level: 3,
                    suffix: LevelSuffix.SPACE,
                    text: "%1)",
                },
            ]);
            const tree = new Formatter().format(abstractNumbering);
            expect(tree).to.deep.equal({
                "w:abstractNum": [
                    {
                        _attr: {
                            "w15:restartNumberingAfterBreak": 0,
                            "w:abstractNumId": 1,
                        },
                    },
                    {
                        "w:multiLevelType": {
                            _attr: {
                                "w:val": "hybridMultilevel",
                            },
                        },
                    },
                    {
                        "w:lvl": [
                            {
                                "w:start": {
                                    _attr: {
                                        "w:val": 1,
                                    },
                                },
                            },
                            {
                                "w:numFmt": {
                                    _attr: {
                                        "w:val": "lowerLetter",
                                    },
                                },
                            },
                            {
                                "w:suff": {
                                    _attr: {
                                        "w:val": "space",
                                    },
                                },
                            },
                            {
                                "w:lvlText": {
                                    _attr: {
                                        "w:val": "%1)",
                                    },
                                },
                            },
                            {
                                "w:lvlJc": {
                                    _attr: {
                                        "w:val": "end",
                                    },
                                },
                            },
                            {
                                _attr: {
                                    "w15:tentative": 1,
                                    "w:ilvl": 3,
                                },
                            },
                        ],
                    },
                ],
            });
            // Expect(tree["w:abstractNum"][2]["w:lvl"]).to.include({ "w:suff": { _attr: { "w:val": "space" } } });
        });

        describe("formatting methods: paragraph properties", () => {
            it("#indent", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                indent: { left: 720 },
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:ind": {
                                                _attr: {
                                                    "w:left": 720,
                                                },
                                            },
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#spacing", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                spacing: { after: 150, before: 50 },
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:spacing": {
                                                _attr: {
                                                    "w:after": 150,
                                                    "w:before": 50,
                                                },
                                            },
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#center", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                alignment: AlignmentType.CENTER,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:jc": {
                                                _attr: {
                                                    "w:val": "center",
                                                },
                                            },
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#left", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                alignment: AlignmentType.LEFT,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
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
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#right", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                alignment: AlignmentType.RIGHT,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:jc": {
                                                _attr: {
                                                    "w:val": "right",
                                                },
                                            },
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#justified", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                alignment: AlignmentType.JUSTIFIED,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:jc": {
                                                _attr: {
                                                    "w:val": "both",
                                                },
                                            },
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#thematicBreak", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                thematicBreak: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
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
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#leftTabStop", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                leftTabStop: 1200,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:tabs": [
                                                {
                                                    "w:tab": {
                                                        _attr: {
                                                            "w:pos": 1200,
                                                            "w:val": "left",
                                                        },
                                                    },
                                                },
                                            ],
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#maxRightTabStop", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                rightTabStop: TabStopPosition.MAX,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:tabs": [
                                                {
                                                    "w:tab": {
                                                        _attr: {
                                                            "w:pos": 9026,
                                                            "w:val": "right",
                                                        },
                                                    },
                                                },
                                            ],
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#keepLines", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                keepLines: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:keepLines": {},
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#keepNext", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            paragraph: {
                                keepNext: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:pPr": [
                                        {
                                            "w:keepNext": {},
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
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
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: { size, sizeComplexScript },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": expected,
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });
            });

            it("#smallCaps", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                smallCaps: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [
                                        {
                                            "w:smallCaps": {},
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#allCaps", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                allCaps: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [
                                        {
                                            "w:caps": {},
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#strike", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                strike: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);

                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [
                                        {
                                            "w:strike": {},
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#doubleStrike", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                doubleStrike: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [
                                        {
                                            "w:dstrike": {},
                                        },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#subScript", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                subScript: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [
                                        { "w:vertAlign": { _attr: { "w:val": "subscript" } } },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#superScript", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                superScript: true,
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [
                                        { "w:vertAlign": { _attr: { "w:val": "superscript" } } },
                                    ],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#font by name", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                font: "Times",
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
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
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });

            it("#font for ascii and eastAsia", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                font: {
                                    ascii: "Times",
                                    eastAsia: "KaiTi",
                                },
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
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
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
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
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: { bold, boldComplexScript },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": expected,
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
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
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: { italics, italicsComplexScript },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": expected,
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
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
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: { highlight, highlightComplexScript },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": expected,
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });
            });

            const shadingTests = [
                {
                    expected: [
                        {
                            "w:shd": {
                                _attr: {
                                    "w:color": "0000FF",
                                    "w:fill": "006622",
                                    "w:val": "diagStripe",
                                },
                            },
                        },
                    ],
                    shading: {
                        color: "0000FF",
                        fill: "006622",
                        type: ShadingType.DIAGONAL_STRIPE,
                    },
                },
                {
                    expected: [
                        {
                            "w:shd": {
                                _attr: {
                                    "w:color": "FF0000",
                                    "w:fill": "00FFFF",
                                    "w:val": "pct10",
                                },
                            },
                        },
                    ],
                    shading: {
                        color: "FF0000",
                        fill: "00FFFF",
                        type: ShadingType.PERCENT_10,
                    },
                },
            ];
            shadingTests.forEach(({ shading, expected }) => {
                it("#shade correctly", () => {
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: { shading },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": expected,
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });
            });

            describe("#underline", () => {
                it("should set underline to 'single' if no arguments are given", () => {
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: {
                                    underline: {},
                                },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": [{ "w:u": { _attr: { "w:val": "single" } } }],
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });

                it("should set the style if given", () => {
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: {
                                    underline: {
                                        type: UnderlineType.DOUBLE,
                                    },
                                },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": [{ "w:u": { _attr: { "w:val": "double" } } }],
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });

                it("should set the style and color if given", () => {
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: {
                                    underline: {
                                        color: "005599",
                                        type: UnderlineType.DOUBLE,
                                    },
                                },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": [
                                            {
                                                "w:u": {
                                                    _attr: {
                                                        "w:color": "005599",
                                                        "w:val": "double",
                                                    },
                                                },
                                            },
                                        ],
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });
            });

            describe("#emphasisMark", () => {
                it("should set emphasisMark to 'dot' if no arguments are given", () => {
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: {
                                    emphasisMark: {},
                                },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": [{ "w:em": { _attr: { "w:val": "dot" } } }],
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });

                it("should set the style if given", () => {
                    const abstractNumbering = new AbstractNumbering(1, [
                        {
                            format: LevelFormat.LOWER_ROMAN,
                            level: 0,
                            style: {
                                run: {
                                    emphasisMark: {
                                        type: EmphasisMarkType.DOT,
                                    },
                                },
                            },
                            text: "%0.",
                        },
                    ]);
                    const tree = new Formatter().format(abstractNumbering);
                    expect(tree).to.deep.equal({
                        "w:abstractNum": [
                            {
                                _attr: {
                                    "w15:restartNumberingAfterBreak": 0,
                                    "w:abstractNumId": 1,
                                },
                            },
                            {
                                "w:multiLevelType": {
                                    _attr: {
                                        "w:val": "hybridMultilevel",
                                    },
                                },
                            },
                            {
                                "w:lvl": [
                                    {
                                        "w:start": {
                                            _attr: {
                                                "w:val": 1,
                                            },
                                        },
                                    },
                                    {
                                        "w:numFmt": {
                                            _attr: {
                                                "w:val": "lowerRoman",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlText": {
                                            _attr: {
                                                "w:val": "%0.",
                                            },
                                        },
                                    },
                                    {
                                        "w:lvlJc": {
                                            _attr: {
                                                "w:val": "start",
                                            },
                                        },
                                    },
                                    {
                                        "w:rPr": [{ "w:em": { _attr: { "w:val": "dot" } } }],
                                    },
                                    {
                                        _attr: {
                                            "w15:tentative": 1,
                                            "w:ilvl": 0,
                                        },
                                    },
                                ],
                            },
                        ],
                    });
                });
            });

            it("#color", () => {
                const abstractNumbering = new AbstractNumbering(1, [
                    {
                        format: LevelFormat.LOWER_ROMAN,
                        level: 0,
                        style: {
                            run: {
                                color: "123456",
                            },
                        },
                        text: "%0.",
                    },
                ]);
                const tree = new Formatter().format(abstractNumbering);
                expect(tree).to.deep.equal({
                    "w:abstractNum": [
                        {
                            _attr: {
                                "w15:restartNumberingAfterBreak": 0,
                                "w:abstractNumId": 1,
                            },
                        },
                        {
                            "w:multiLevelType": {
                                _attr: {
                                    "w:val": "hybridMultilevel",
                                },
                            },
                        },
                        {
                            "w:lvl": [
                                {
                                    "w:start": {
                                        _attr: {
                                            "w:val": 1,
                                        },
                                    },
                                },
                                {
                                    "w:numFmt": {
                                        _attr: {
                                            "w:val": "lowerRoman",
                                        },
                                    },
                                },
                                {
                                    "w:lvlText": {
                                        _attr: {
                                            "w:val": "%0.",
                                        },
                                    },
                                },
                                {
                                    "w:lvlJc": {
                                        _attr: {
                                            "w:val": "start",
                                        },
                                    },
                                },
                                {
                                    "w:rPr": [{ "w:color": { _attr: { "w:val": "123456" } } }],
                                },
                                {
                                    _attr: {
                                        "w15:tentative": 1,
                                        "w:ilvl": 0,
                                    },
                                },
                            ],
                        },
                    ],
                });
            });
        });
    });
});
