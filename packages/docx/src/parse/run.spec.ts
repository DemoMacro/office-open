import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxParseContext } from "./context";
import { parseRun } from "./run";
import type { TextRunJson } from "./types";

const mockCtx: DocxParseContext = {
    zip: new Map(),
    hyperlinks: new Map(),
    media: new Map(),
    documentRels: new Map(),
};

function asTextRun(result: ReturnType<typeof parseRun>): TextRunJson {
    return result as TextRunJson;
}

describe("parseRun", () => {
    it("should parse plain text run", () => {
        const run: Element = {
            elements: [{ name: "w:t", text: "Hello" }],
        };
        const result = parseRun(run, mockCtx);
        expect(result?.$type).toBe("textRun");
        expect(asTextRun(result).text).toBe("Hello");
    });

    it("should parse bold property", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:b" }] },
                { name: "w:t", text: "Bold" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.bold).toBe(true);
    });

    it("should parse italic property", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:i" }] },
                { name: "w:t", text: "Italic" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.italics).toBe(true);
    });

    it("should parse underline", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:u", attributes: { "w:val": "single" } }] },
                { name: "w:t", text: "Under" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.underline).toEqual({ type: "single" });
    });

    it("should parse font size (half-points)", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:sz", attributes: { "w:val": "28" } }] },
                { name: "w:t", text: "Big" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.size).toBe(28);
    });

    it("should parse color", () => {
        const run: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:color", attributes: { "w:val": "FF0000" } }],
                },
                { name: "w:t", text: "Red" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.color).toBe("FF0000");
    });

    it("should parse font family", () => {
        const run: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:rFonts", attributes: { "w:ascii": "Calibri" } }],
                },
                { name: "w:t", text: "Font" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.font).toBe("Calibri");
    });

    it("should parse strike", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:strike" }] },
                { name: "w:t", text: "Struck" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.strike).toBe(true);
    });

    it("should parse subscript/superscript", () => {
        const subRun: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:vertAlign", attributes: { "w:val": "subscript" } }],
                },
                { name: "w:t", text: "sub" },
            ],
        };
        expect(asTextRun(parseRun(subRun, mockCtx)).subScript).toBe(true);

        const superRun: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:vertAlign", attributes: { "w:val": "superscript" } }],
                },
                { name: "w:t", text: "super" },
            ],
        };
        expect(asTextRun(parseRun(superRun, mockCtx)).superScript).toBe(true);
    });

    it("should parse page break", () => {
        const run: Element = {
            elements: [{ name: "w:br", attributes: { "w:type": "page" } }],
        };
        const result = parseRun(run, mockCtx);
        expect(result?.$type).toBe("pageBreak");
    });

    it("should parse tab", () => {
        const run: Element = {
            elements: [{ name: "w:tab" }],
        };
        const result = parseRun(run, mockCtx);
        expect(result?.$type).toBe("tab");
    });

    it("should return undefined for empty run without properties", () => {
        const run: Element = { elements: [] };
        expect(parseRun(run, mockCtx)).toBeUndefined();
    });

    it("should parse highlight", () => {
        const run: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:highlight", attributes: { "w:val": "yellow" } }],
                },
                { name: "w:t", text: "Hi" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.highlight).toBe("yellow");
    });

    it("should parse character spacing", () => {
        const run: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:spacing", attributes: { "w:val": "120" } }],
                },
                { name: "w:t", text: "Spaced" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.characterSpacing).toBe(120);
    });

    it("should parse shading", () => {
        const run: Element = {
            elements: [
                {
                    name: "w:rPr",
                    elements: [{ name: "w:shd", attributes: { "w:fill": "FFFF00" } }],
                },
                { name: "w:t", text: "Shaded" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.shading).toEqual({ fill: "FFFF00" });
    });

    it("should parse kern", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:kern", attributes: { "w:val": "24" } }] },
                { name: "w:t", text: "Kerned" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.kern).toBe(24);
    });

    it("should parse position", () => {
        const run: Element = {
            elements: [
                { name: "w:rPr", elements: [{ name: "w:position", attributes: { "w:val": "6" } }] },
                { name: "w:t", text: "Pos" },
            ],
        };
        const result = asTextRun(parseRun(run, mockCtx));
        expect(result.position).toBe("6");
    });
});
