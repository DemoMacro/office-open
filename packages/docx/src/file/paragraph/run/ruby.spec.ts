import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createRuby } from "./ruby";

describe("createRuby", () => {
    it("should create ruby with default options", () => {
        const tree = new Formatter().format(
            createRuby({
                text: "かな",
                base: "漢字",
            }),
        );
        expect(tree).to.deep.equal({
            "w:ruby": [
                {
                    "w:rubyPr": [
                        { "w:rubyAlign": { _attr: { "w:val": "center" } } },
                        { "w:hps": { _attr: { "w:val": 20 } } },
                        { "w:hpsRaise": { _attr: { "w:val": 20 } } },
                        { "w:hpsBaseText": { _attr: { "w:val": 40 } } },
                        { "w:lid": { _attr: { "w:val": "ja-JP" } } },
                    ],
                },
                {
                    "w:rt": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "かな"],
                                },
                            ],
                        },
                    ],
                },
                {
                    "w:rubyBase": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "漢字"],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create ruby with custom alignment", () => {
        const tree = new Formatter().format(
            createRuby({
                text: "pinyin",
                base: "汉字",
                alignment: "LEFT",
            }),
        );
        expect(tree).to.deep.equal({
            "w:ruby": [
                {
                    "w:rubyPr": [
                        { "w:rubyAlign": { _attr: { "w:val": "left" } } },
                        { "w:hps": { _attr: { "w:val": 20 } } },
                        { "w:hpsRaise": { _attr: { "w:val": 20 } } },
                        { "w:hpsBaseText": { _attr: { "w:val": 40 } } },
                        { "w:lid": { _attr: { "w:val": "ja-JP" } } },
                    ],
                },
                {
                    "w:rt": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "pinyin"],
                                },
                            ],
                        },
                    ],
                },
                {
                    "w:rubyBase": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "汉字"],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create ruby with custom font sizes", () => {
        const tree = new Formatter().format(
            createRuby({
                text: "ルビ",
                base: "本文",
                fontSize: 16,
                raise: 24,
                baseFontSize: 32,
            }),
        );
        expect(tree).to.deep.equal({
            "w:ruby": [
                {
                    "w:rubyPr": [
                        { "w:rubyAlign": { _attr: { "w:val": "center" } } },
                        { "w:hps": { _attr: { "w:val": 16 } } },
                        { "w:hpsRaise": { _attr: { "w:val": 24 } } },
                        { "w:hpsBaseText": { _attr: { "w:val": 32 } } },
                        { "w:lid": { _attr: { "w:val": "ja-JP" } } },
                    ],
                },
                {
                    "w:rt": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "ルビ"],
                                },
                            ],
                        },
                    ],
                },
                {
                    "w:rubyBase": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "本文"],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create ruby with custom language", () => {
        const tree = new Formatter().format(
            createRuby({
                text: "zhuyin",
                base: "注音",
                languageId: "zh-TW",
            }),
        );
        expect(tree).to.deep.equal({
            "w:ruby": [
                {
                    "w:rubyPr": [
                        { "w:rubyAlign": { _attr: { "w:val": "center" } } },
                        { "w:hps": { _attr: { "w:val": 20 } } },
                        { "w:hpsRaise": { _attr: { "w:val": 20 } } },
                        { "w:hpsBaseText": { _attr: { "w:val": 40 } } },
                        { "w:lid": { _attr: { "w:val": "zh-TW" } } },
                    ],
                },
                {
                    "w:rt": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "zhuyin"],
                                },
                            ],
                        },
                    ],
                },
                {
                    "w:rubyBase": [
                        {
                            "w:r": [
                                {
                                    "w:t": [{ _attr: { "xml:space": "preserve" } }, "注音"],
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });
});
