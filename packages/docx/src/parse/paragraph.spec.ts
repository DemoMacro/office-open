import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxParseContext } from "./context";
import { parseParagraph } from "./paragraph";
import type { TextRunJson, BookmarkJson } from "./types";

const mockCtx: DocxParseContext = {
    zip: new Map(),
    hyperlinks: new Map(),
    media: new Map(),
    documentRels: new Map(),
};

describe("parseParagraph", () => {
    it("should parse plain text paragraph", () => {
        const p: Element = {
            elements: [{ name: "w:r", elements: [{ name: "w:t", text: "Hello" }] }],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.$type).toBe("paragraph");
        expect(result.text).toBe("Hello");
    });

    it("should detect heading from style", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:pPr",
                    elements: [{ name: "w:pStyle", attributes: { "w:val": "HEADING_1" } }],
                },
                { name: "w:r", elements: [{ name: "w:t", text: "Title" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.heading).toBe("HEADING_1");
    });

    it("should parse paragraph alignment", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:pPr",
                    elements: [{ name: "w:jc", attributes: { "w:val": "center" } }],
                },
                { name: "w:r", elements: [{ name: "w:t", text: "Centered" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.alignment).toBe("center");
    });

    it("should parse spacing", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:pPr",
                    elements: [
                        { name: "w:spacing", attributes: { "w:before": "200", "w:after": "400" } },
                    ],
                },
                { name: "w:r", elements: [{ name: "w:t", text: "Spaced" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.spacing?.before).toBe(200);
        expect(result.spacing?.after).toBe(400);
    });

    it("should parse indentation", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:pPr",
                    elements: [
                        { name: "w:ind", attributes: { "w:left": "720", "w:firstLine": "480" } },
                    ],
                },
                { name: "w:r", elements: [{ name: "w:t", text: "Indented" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.indent?.left).toBe(720);
        expect(result.indent?.firstLine).toBe(480);
    });

    it("should parse numbering", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:pPr",
                    elements: [
                        {
                            name: "w:numPr",
                            elements: [
                                { name: "w:numId", attributes: { "w:val": "1" } },
                                { name: "w:ilvl", attributes: { "w:val": "0" } },
                            ],
                        },
                    ],
                },
                { name: "w:r", elements: [{ name: "w:t", text: "List item" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.numbering?.reference).toBe("1");
        expect(result.numbering?.level).toBe(0);
    });

    it("should preserve rich text runs as children", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:r",
                    elements: [
                        { name: "w:rPr", elements: [{ name: "w:b" }] },
                        { name: "w:t", text: "Bold" },
                    ],
                },
                {
                    name: "w:r",
                    elements: [
                        { name: "w:rPr", elements: [{ name: "w:i" }] },
                        { name: "w:t", text: "Italic" },
                    ],
                },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.children).toHaveLength(2);
        expect((result.children![0] as TextRunJson).bold).toBe(true);
        expect((result.children![1] as TextRunJson).italics).toBe(true);
    });

    it("should parse bookmark", () => {
        const p: Element = {
            elements: [
                { name: "w:bookmarkStart", attributes: { "w:name": "_Toc123" } },
                { name: "w:r", elements: [{ name: "w:t", text: "text" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.children).toBeDefined();
        expect((result.children![0] as BookmarkJson).$type).toBe("bookmark");
        expect((result.children![0] as BookmarkJson).name).toBe("_Toc123");
    });

    it("should parse thematic break from bottom border", () => {
        const p: Element = {
            elements: [
                {
                    name: "w:pPr",
                    elements: [
                        {
                            name: "w:pBdr",
                            elements: [{ name: "w:bottom", attributes: { "w:val": "single" } }],
                        },
                    ],
                },
                { name: "w:r", elements: [{ name: "w:t", text: "text" }] },
            ],
        };
        const result = parseParagraph(p, mockCtx);
        expect(result.thematicBreak).toBe(true);
    });
});
