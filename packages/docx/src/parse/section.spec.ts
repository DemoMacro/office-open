import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseSectionProperties } from "./section";

function get<T = unknown>(obj: Record<string, unknown> | undefined, key: string): T {
    return obj?.[key] as T;
}

describe("parseSectionProperties", () => {
    it("should parse page size", () => {
        const sectPr: Element = {
            elements: [
                {
                    name: "w:pgSz",
                    attributes: { "w:w": "11906", "w:h": "16838", "w:orient": "portrait" },
                },
            ],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(get<Record<string, unknown>>(result, "page")?.size).toEqual({
            width: 11906,
            height: 16838,
            orientation: "portrait",
        });
    });

    it("should parse page margins", () => {
        const sectPr: Element = {
            elements: [
                {
                    name: "w:pgMar",
                    attributes: {
                        "w:top": "1440",
                        "w:right": "1440",
                        "w:bottom": "1440",
                        "w:left": "1440",
                    },
                },
            ],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(get<Record<string, unknown>>(result, "page")?.margin).toEqual({
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440,
        });
    });

    it("should parse section type", () => {
        const sectPr: Element = {
            elements: [{ name: "w:type", attributes: { "w:val": "nextPage" } }],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(result.type).toBe("nextPage");
    });

    it("should parse title page", () => {
        const sectPr: Element = {
            elements: [{ name: "w:titlePg" }],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(result.titlePage).toBe(true);
    });

    it("should parse page number start", () => {
        const sectPr: Element = {
            elements: [{ name: "w:pgNumType", attributes: { "w:start": "3" } }],
        };
        const result = parseSectionProperties(sectPr)!;
        const page = get<Record<string, unknown>>(result, "page");
        const pageNumbers = get<Record<string, unknown>>(page, "pageNumbers");
        expect(pageNumbers?.start).toBe(3);
    });

    it("should parse columns", () => {
        const sectPr: Element = {
            elements: [{ name: "w:cols", attributes: { "w:num": "2", "w:space": "708" } }],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(get<Record<string, unknown>>(result, "column")).toEqual({ count: 2, space: 708 });
    });

    it("should parse vertical align", () => {
        const sectPr: Element = {
            elements: [{ name: "w:vAlign", attributes: { "w:val": "center" } }],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(result.verticalAlign).toBe("center");
    });

    it("should parse header/footer references", () => {
        const sectPr: Element = {
            elements: [
                { name: "w:headerReference", attributes: { "w:type": "default", "r:id": "rId1" } },
                { name: "w:footerReference", attributes: { "w:type": "default", "r:id": "rId2" } },
            ],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(get<Record<string, string>>(result, "headerRefs")?.default).toBe("rId1");
        expect(get<Record<string, string>>(result, "footerRefs")?.default).toBe("rId2");
    });

    it("should return undefined for empty sectPr", () => {
        const result = parseSectionProperties({ elements: [] });
        expect(result).toBeUndefined();
    });

    it("should parse line numbers", () => {
        const sectPr: Element = {
            elements: [
                {
                    name: "w:lnNumType",
                    attributes: { "w:countBy": "1", "w:restart": "newPage", "w:start": "1" },
                },
            ],
        };
        const result = parseSectionProperties(sectPr)!;
        expect(get<Record<string, unknown>>(result, "lineNumbers")).toEqual({
            countBy: 1,
            restart: "newPage",
            start: 1,
        });
    });

    it("should parse page borders", () => {
        const sectPr: Element = {
            elements: [
                {
                    name: "w:pgBorders",
                    elements: [
                        {
                            name: "w:top",
                            attributes: { "w:val": "single", "w:sz": "4", "w:color": "000000" },
                        },
                    ],
                },
            ],
        };
        const result = parseSectionProperties(sectPr)!;
        const page = get<Record<string, unknown>>(result, "page");
        const borders = get<Record<string, unknown>>(page, "borders");
        expect(borders?.top).toEqual({
            style: "single",
            size: 4,
            color: "000000",
        });
    });
});
