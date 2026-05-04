import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { PptxParseContext } from "./context";
import { parseTextBody, parsePptxParagraph } from "./text-body";
import type { RunJson } from "./types";

const mockCtx: PptxParseContext = {
    zip: new Map(),
    slidePaths: [],
};

describe("parseTextBody", () => {
    it("should parse body properties", () => {
        const txBody: Element = {
            elements: [
                {
                    name: "a:bodyPr",
                    attributes: { vert: "vert270", anchor: "ctr" },
                },
                {
                    name: "a:p",
                    elements: [{ name: "a:r", elements: [{ name: "a:t", text: "Hi" }] }],
                },
            ],
        };
        const result = parseTextBody(txBody, mockCtx);
        expect(result.vertical).toBe("vert270");
        expect(result.anchor).toBe("ctr");
        expect(result.paragraphs).toHaveLength(1);
    });

    it("should parse autoFit", () => {
        const txBody: Element = {
            elements: [{ name: "a:bodyPr", elements: [{ name: "a:spAutoFit" }] }],
        };
        const result = parseTextBody(txBody, mockCtx);
        expect(result.autoFit).toBe("shape");
    });

    it("should parse margins", () => {
        const txBody: Element = {
            elements: [
                {
                    name: "a:bodyPr",
                    attributes: { lIns: "91440", tIns: "45720", rIns: "91440", bIns: "45720" },
                },
            ],
        };
        const result = parseTextBody(txBody, mockCtx);
        expect(result.margins).toEqual({
            left: 91440,
            top: 45720,
            right: 91440,
            bottom: 45720,
        });
    });
});

describe("parsePptxParagraph", () => {
    it("should parse paragraph with text", () => {
        const p: Element = {
            elements: [{ name: "a:r", elements: [{ name: "a:t", text: "Hello" }] }],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect(result.text).toBe("Hello");
    });

    it("should parse alignment", () => {
        const p: Element = {
            elements: [
                { name: "a:pPr", attributes: { algn: "ctr" } },
                { name: "a:r", elements: [{ name: "a:t", text: "C" }] },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect(result.alignment).toBe("ctr");
    });

    it("should parse bullet", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:pPr",
                    elements: [{ name: "a:buChar", attributes: { char: "•" } }],
                },
                { name: "a:r", elements: [{ name: "a:t", text: "Bullet" }] },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect(result.bullet).toEqual({ type: "•", level: 0 });
    });

    it("should parse line break", () => {
        const p: Element = {
            elements: [
                { name: "a:r", elements: [{ name: "a:t", text: "A" }] },
                { name: "a:br" },
                { name: "a:r", elements: [{ name: "a:t", text: "B" }] },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect(result.children).toHaveLength(3);
        expect(result.children![1].text).toBe("\n");
    });
});

describe("parsePptxRun (via parsePptxParagraph)", () => {
    it("should parse font size", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:r",
                    elements: [
                        { name: "a:rPr", attributes: { sz: "2400" } },
                        { name: "a:t", text: "Big" },
                    ],
                },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect((result.children![0] as RunJson).fontSize).toBe(2400);
    });

    it("should parse bold", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:r",
                    elements: [
                        { name: "a:rPr", attributes: { b: "1" } },
                        { name: "a:t", text: "B" },
                    ],
                },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect((result.children![0] as RunJson).bold).toBe(true);
    });

    it("should parse italic", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:r",
                    elements: [
                        { name: "a:rPr", attributes: { i: "1" } },
                        { name: "a:t", text: "I" },
                    ],
                },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect((result.children![0] as RunJson).italic).toBe(true);
    });

    it("should parse underline", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:r",
                    elements: [
                        { name: "a:rPr", attributes: { u: "sng" } },
                        { name: "a:t", text: "U" },
                    ],
                },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect((result.children![0] as RunJson).underline).toBe("sng");
    });

    it("should parse font family", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:r",
                    elements: [
                        {
                            name: "a:rPr",
                            elements: [{ name: "a:latin", attributes: { typeface: "Arial" } }],
                        },
                        { name: "a:t", text: "F" },
                    ],
                },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect((result.children![0] as RunJson).font).toBe("Arial");
    });

    it("should parse solid fill color", () => {
        const p: Element = {
            elements: [
                {
                    name: "a:r",
                    elements: [
                        {
                            name: "a:rPr",
                            elements: [
                                {
                                    name: "a:solidFill",
                                    elements: [
                                        { name: "a:srgbClr", attributes: { val: "FF0000" } },
                                    ],
                                },
                            ],
                        },
                        { name: "a:t", text: "R" },
                    ],
                },
            ],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect((result.children![0] as RunJson).fill).toEqual({ type: "solid", color: "FF0000" });
    });

    it("should skip empty runs without properties", () => {
        const p: Element = {
            elements: [{ name: "a:r", elements: [] }],
        };
        const result = parsePptxParagraph(p, mockCtx);
        expect(result.children).toBeUndefined();
    });
});
