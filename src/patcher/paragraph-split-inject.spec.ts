import { describe, expect, it } from "vite-plus/test";

import { findRunElementIndexWithToken, splitRunElement } from "./paragraph-split-inject";

describe("paragraph-split-inject", () => {
    describe("findRunElementIndexWithToken", () => {
        it("should find the index of a run element with a token", () => {
            const output = findRunElementIndexWithToken(
                {
                    elements: [
                        {
                            elements: [
                                {
                                    name: "w:t",
                                    type: "element",
                                    elements: [
                                        {
                                            type: "text",
                                            text: "hello world",
                                        },
                                    ],
                                },
                            ],
                            name: "w:r",
                            type: "element",
                        },
                    ],
                    name: "w:p",
                    type: "element",
                },
                "hello",
            );
            expect(output).to.deep.equal(0);
        });

        it("should throw an exception when ran with empty elements", () => {
            expect(() =>
                findRunElementIndexWithToken(
                    {
                        name: "w:p",
                        type: "element",
                    },
                    "hello",
                ),
            ).to.throw();
        });

        it("should throw an exception when ran with empty elements", () => {
            expect(() =>
                findRunElementIndexWithToken(
                    {
                        elements: [
                            {
                                name: "w:r",
                                type: "element",
                            },
                        ],
                        name: "w:p",
                        type: "element",
                    },
                    "hello",
                ),
            ).to.throw();
        });

        it("should throw an exception when ran with empty elements", () => {
            expect(() =>
                findRunElementIndexWithToken(
                    {
                        elements: [
                            {
                                elements: [
                                    {
                                        name: "w:t",
                                        type: "element",
                                    },
                                ],
                                name: "w:r",
                                type: "element",
                            },
                        ],
                        name: "w:p",
                        type: "element",
                    },
                    "hello",
                ),
            ).to.throw();
        });

        it("should continue if text run doesn't have text", () => {
            expect(() =>
                findRunElementIndexWithToken(
                    {
                        elements: [
                            {
                                elements: [
                                    {
                                        name: "w:t",
                                        type: "element",
                                    },
                                ],
                                name: "w:r",
                                type: "element",
                            },
                        ],
                        name: "w:p",
                        type: "element",
                    },
                    "hello",
                ),
            ).to.throw();
        });

        it("should continue if text run doesn't have text", () => {
            expect(() =>
                findRunElementIndexWithToken(
                    {
                        elements: [
                            {
                                elements: [
                                    {
                                        name: "w:t",
                                        type: "element",
                                        elements: [
                                            {
                                                type: "text",
                                            },
                                        ],
                                    },
                                ],
                                name: "w:r",
                                type: "element",
                            },
                        ],
                        name: "w:p",
                        type: "element",
                    },
                    "hello",
                ),
            ).to.throw();
        });
    });

    describe("splitRunElement", () => {
        it("should split a run element", () => {
            const output = splitRunElement(
                {
                    elements: [
                        {
                            elements: [
                                {
                                    type: "text",
                                    text: "hello*world",
                                },
                            ],
                            name: "w:t",
                            type: "element",
                        },
                        {
                            name: "w:x",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
                "*",
            );

            expect(output).to.deep.equal({
                left: {
                    elements: [
                        {
                            attributes: {
                                "xml:space": "preserve",
                            },
                            elements: [
                                {
                                    text: "hello",
                                    type: "text",
                                },
                            ],
                            name: "w:t",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
                right: {
                    elements: [
                        {
                            attributes: {
                                "xml:space": "preserve",
                            },
                            elements: [
                                {
                                    text: "world",
                                    type: "text",
                                },
                            ],
                            name: "w:t",
                            type: "element",
                        },
                        {
                            name: "w:x",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
            });
        });

        it("should try to split even if elements is empty for text", () => {
            const output = splitRunElement(
                {
                    elements: [
                        {
                            name: "w:t",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
                "*",
            );

            // When the token is not found in the text, splitIndex remains -1
            // So left gets nothing and right gets all elements
            expect(output).to.deep.equal({
                left: {
                    elements: [],
                    name: "w:r",
                    type: "element",
                },
                right: {
                    elements: [
                        {
                            attributes: {
                                "xml:space": "preserve",
                            },
                            elements: [],
                            name: "w:t",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
            });
        });

        it("should return empty elements", () => {
            const output = splitRunElement(
                {
                    name: "w:r",
                    type: "element",
                },
                "*",
            );

            expect(output).to.deep.equal({
                left: {
                    elements: [],
                    name: "w:r",
                    type: "element",
                },
                right: {
                    elements: [],
                    name: "w:r",
                    type: "element",
                },
            });
        });

        it("should put all content on the right when token is not found", () => {
            const output = splitRunElement(
                {
                    elements: [
                        {
                            elements: [
                                {
                                    type: "text",
                                    text: "hello world",
                                },
                            ],
                            name: "w:t",
                            type: "element",
                        },
                        {
                            name: "w:x",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
                "*",
            );

            // When the token is not found, splitIndex remains -1
            // So left gets nothing and right gets all elements
            expect(output).to.deep.equal({
                left: {
                    elements: [],
                    name: "w:r",
                    type: "element",
                },
                right: {
                    elements: [
                        {
                            attributes: {
                                "xml:space": "preserve",
                            },
                            elements: [
                                {
                                    text: "hello world",
                                    type: "text",
                                },
                            ],
                            name: "w:t",
                            type: "element",
                        },
                        {
                            name: "w:x",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
            });
        });

        it("should create an empty end element if it is at the end", () => {
            const output = splitRunElement(
                {
                    elements: [
                        {
                            elements: [
                                {
                                    type: "element",
                                    name: "w:rFonts",
                                    attributes: { "w:eastAsia": "Times New Roman" },
                                },
                                { type: "element", name: "w:kern", attributes: { "w:val": "0" } },
                                { type: "element", name: "w:sz", attributes: { "w:val": "20" } },
                                {
                                    type: "element",
                                    name: "w:lang",
                                    attributes: {
                                        "w:val": "en-US",
                                        "w:eastAsia": "en-US",
                                        "w:bidi": "ar-SA",
                                    },
                                },
                            ],
                            name: "w:rPr",
                            type: "element",
                        },
                        {
                            attributes: { "xml:space": "preserve" },
                            elements: [],
                            name: "w:t",
                            type: "element",
                        },
                        { name: "w:br", type: "element" },
                        { elements: [{ type: "text", text: "ɵ" }], name: "w:t", type: "element" },
                    ],
                    name: "w:r",
                    type: "element",
                },
                "ɵ",
            );

            expect(output).to.deep.equal({
                left: {
                    elements: [
                        {
                            elements: [
                                {
                                    type: "element",
                                    name: "w:rFonts",
                                    attributes: { "w:eastAsia": "Times New Roman" },
                                },
                                { type: "element", name: "w:kern", attributes: { "w:val": "0" } },
                                { type: "element", name: "w:sz", attributes: { "w:val": "20" } },
                                {
                                    type: "element",
                                    name: "w:lang",
                                    attributes: {
                                        "w:val": "en-US",
                                        "w:eastAsia": "en-US",
                                        "w:bidi": "ar-SA",
                                    },
                                },
                            ],
                            name: "w:rPr",
                            type: "element",
                        },
                        {
                            attributes: { "xml:space": "preserve" },
                            elements: [],
                            name: "w:t",
                            type: "element",
                        },
                        { name: "w:br", type: "element" },
                        {
                            attributes: { "xml:space": "preserve" },
                            elements: [],
                            name: "w:t",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
                right: {
                    elements: [
                        {
                            attributes: { "xml:space": "preserve" },
                            elements: [],
                            name: "w:t",
                            type: "element",
                        },
                    ],
                    name: "w:r",
                    type: "element",
                },
            });
        });
    });
});
