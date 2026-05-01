import { describe, expect, it } from "vite-plus/test";

import { replaceTokenInParagraphElement } from "./paragraph-token-replacer";

describe("paragraph-token-replacer", () => {
    describe("replaceTokenInParagraphElement", () => {
        it("should replace token in paragraph", () => {
            const output = replaceTokenInParagraphElement({
                originalText: "hello",
                paragraphElement: {
                    elements: [
                        {
                            elements: [
                                {
                                    name: "w:t",
                                    elements: [
                                        {
                                            type: "text",
                                            text: "hello",
                                        },
                                    ],
                                },
                            ],
                            name: "w:r",
                        },
                    ],
                    name: "w:p",
                },
                renderedParagraph: {
                    index: 0,
                    pathToParagraph: [0],
                    runs: [
                        {
                            end: 4,
                            index: 0,
                            parts: [
                                {
                                    end: 4,
                                    index: 0,
                                    start: 0,
                                    text: "hello",
                                },
                            ],
                            start: 0,
                            text: "hello",
                        },
                    ],
                    text: "hello",
                },
                replacementText: "world",
            });

            expect(output).to.deep.equal({
                elements: [
                    {
                        elements: [
                            {
                                elements: [
                                    {
                                        text: "world",
                                        type: "text",
                                    },
                                ],
                                name: "w:t",
                            },
                        ],
                        name: "w:r",
                    },
                ],
                name: "w:p",
            });
        });

        it("should handle case where it cannot find any text to replace", () => {
            const output = replaceTokenInParagraphElement({
                originalText: "{{name}}",
                paragraphElement: {
                    attributes: {
                        "w14:paraId": "2499FE9F",
                        "w14:textId": "27B4FBC2",
                        "w:rsidP": "007B52ED",
                        "w:rsidR": "00B51233",
                        "w:rsidRDefault": "007B52ED",
                    },
                    elements: [
                        {
                            elements: [
                                {
                                    type: "element",
                                    name: "w:pStyle",
                                    attributes: { "w:val": "Title" },
                                },
                            ],
                            name: "w:pPr",
                            type: "element",
                        },
                        {
                            elements: [
                                {
                                    type: "element",
                                    name: "w:t",
                                    attributes: { "xml:space": "preserve" },
                                    elements: [{ type: "text", text: "Hello " }],
                                },
                            ],
                            name: "w:r",
                            type: "element",
                        },
                        {
                            attributes: { "w:rsidR": "007F116B" },
                            elements: [
                                {
                                    type: "element",
                                    name: "w:t",
                                    attributes: { "xml:space": "preserve" },
                                    elements: [{ type: "text", text: "{{name}} " }],
                                },
                            ],
                            name: "w:r",
                            type: "element",
                        },
                        {
                            elements: [
                                {
                                    type: "element",
                                    name: "w:t",
                                    elements: [{ type: "text", text: "World" }],
                                },
                            ],
                            name: "w:r",
                            type: "element",
                        },
                    ],
                    name: "w:p",
                },
                renderedParagraph: {
                    index: 0,
                    pathToParagraph: [0, 1, 0, 0],
                    runs: [
                        {
                            end: 5,
                            index: 1,
                            parts: [{ text: "Hello ", index: 0, start: 0, end: 5 }],
                            start: 0,
                            text: "Hello ",
                        },
                        {
                            end: 14,
                            index: 2,
                            parts: [{ text: "{{name}} ", index: 0, start: 6, end: 14 }],
                            start: 6,
                            text: "{{name}} ",
                        },
                        {
                            end: 19,
                            index: 3,
                            parts: [{ text: "World", index: 0, start: 15, end: 19 }],
                            start: 15,
                            text: "World",
                        },
                    ],
                    text: "Hello {{name}} World",
                },
                replacementText: "John",
            });

            expect(output).to.deep.equal({
                attributes: {
                    "w14:paraId": "2499FE9F",
                    "w14:textId": "27B4FBC2",
                    "w:rsidP": "007B52ED",
                    "w:rsidR": "00B51233",
                    "w:rsidRDefault": "007B52ED",
                },
                elements: [
                    {
                        elements: [
                            {
                                attributes: {
                                    "w:val": "Title",
                                },
                                name: "w:pStyle",
                                type: "element",
                            },
                        ],
                        name: "w:pPr",
                        type: "element",
                    },
                    {
                        elements: [
                            {
                                attributes: {
                                    "xml:space": "preserve",
                                },
                                elements: [
                                    {
                                        text: "Hello ",
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
                    {
                        attributes: {
                            "w:rsidR": "007F116B",
                        },
                        elements: [
                            {
                                attributes: {
                                    "xml:space": "preserve",
                                },
                                elements: [
                                    {
                                        text: "John ",
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
                    {
                        elements: [
                            {
                                attributes: {
                                    "xml:space": "preserve",
                                },
                                elements: [
                                    {
                                        text: "World",
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
                ],
                name: "w:p",
            });
        });

        it("should skip part when partToReplace is empty", () => {
            const output = replaceTokenInParagraphElement({
                originalText: "hello",
                paragraphElement: {
                    elements: [
                        {
                            elements: [
                                {
                                    name: "w:t",
                                    elements: [{ type: "text", text: "" }],
                                },
                            ],
                            name: "w:r",
                        },
                        {
                            elements: [
                                {
                                    name: "w:t",
                                    elements: [{ type: "text", text: "hello" }],
                                },
                            ],
                            name: "w:r",
                        },
                    ],
                    name: "w:p",
                },
                renderedParagraph: {
                    index: 0,
                    pathToParagraph: [0],
                    runs: [
                        {
                            end: 0,
                            index: 0,
                            parts: [
                                {
                                    end: 0,
                                    index: 0,
                                    start: 0,
                                    text: "",
                                },
                            ],
                            start: 0,
                            text: "",
                        },
                        {
                            end: 5,
                            index: 1,
                            parts: [
                                {
                                    end: 5,
                                    index: 0,
                                    start: 0,
                                    text: "hello",
                                },
                            ],
                            start: 0,
                            text: "hello",
                        },
                    ],
                    text: "hello",
                },
                replacementText: "world",
            });

            expect(output).to.deep.equal({
                elements: [
                    {
                        elements: [
                            {
                                elements: [{ text: "", type: "text" }],
                                name: "w:t",
                            },
                        ],
                        name: "w:r",
                    },
                    {
                        elements: [
                            {
                                elements: [{ text: "world", type: "text" }],
                                name: "w:t",
                            },
                        ],
                        name: "w:r",
                    },
                ],
                name: "w:p",
            });
        });

        // Try to fill rest of test coverage
        // It("should replace token in paragraph", () => {
        //     Const output = replaceTokenInParagraphElement({
        //         ParagraphElement: {
        //             Name: "w:p",
        //             Elements: [
        //                 {
        //                     Name: "w:r",
        //                     Elements: [
        //                         {
        //                             Name: "w:t",
        //                             Elements: [
        //                                 {
        //                                     Type: "text",
        //                                     Text: "test ",
        //                                 },
        //                             ],
        //                         },
        //                         {
        //                             Name: "w:t",
        //                             Elements: [
        //                                 {
        //                                     Type: "text",
        //                                     Text: " hello ",
        //                                 },
        //                             ],
        //                         },
        //                     ],
        //                 },
        //             ],
        //         },
        //         RenderedParagraph: {
        //             Index: 0,
        //             Path: [0],
        //             Runs: [
        //                 {
        //                     End: 4,
        //                     Index: 0,
        //                     Parts: [
        //                         {
        //                             End: 4,
        //                             Index: 0,
        //                             Start: 0,
        //                             Text: "test ",
        //                         },
        //                     ],
        //                     Start: 0,
        //                     Text: "test ",
        //                 },
        //                 {
        //                     End: 10,
        //                     Index: 0,
        //                     Parts: [
        //                         {
        //                             End: 10,
        //                             Index: 0,
        //                             Start: 5,
        //                             Text: "hello ",
        //                         },
        //                     ],
        //                     Start: 5,
        //                     Text: "hello ",
        //                 },
        //             ],
        //             Text: "test hello ",
        //         },
        //         OriginalText: "hello",
        //         ReplacementText: "world",
        //     });

        //     Expect(output).to.deep.equal({
        //         Elements: [
        //             {
        //                 Elements: [
        //                     {
        //                         Elements: [
        //                             {
        //                                 Text: "test world ",
        //                                 Type: "text",
        //                             },
        //                         ],
        //                         Name: "w:t",
        //                     },
        //                 ],
        //                 Name: "w:r",
        //             },
        //         ],
        //         Name: "w:p",
        //     });
        // });
    });
});
