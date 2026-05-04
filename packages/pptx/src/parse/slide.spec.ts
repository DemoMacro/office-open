import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { PptxParseContext } from "./context";
import { parseSlide } from "./slide";

const mockCtx: PptxParseContext = {
    zip: new Map(),
    slidePaths: ["ppt/slides/slide1.xml"],
};

describe("parseSlide", () => {
    it("should parse slide with shapes", () => {
        const slideXml: Element = {
            elements: [
                {
                    name: "p:cSld",
                    elements: [
                        {
                            name: "p:spTree",
                            elements: [
                                {
                                    name: "p:sp",
                                    elements: [
                                        {
                                            name: "p:nvSpPr",
                                            elements: [
                                                {
                                                    name: "p:cNvPr",
                                                    attributes: { id: "2", name: "Shape 1" },
                                                },
                                            ],
                                        },
                                        { name: "p:spPr", elements: [] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseSlide(slideXml, mockCtx, "ppt/slides/slide1.xml");
        expect(result.children).toHaveLength(1);
        expect(result.children[0].$type).toBe("shape");
    });

    it("should parse slide background", () => {
        const slideXml: Element = {
            elements: [
                {
                    name: "p:cSld",
                    elements: [
                        {
                            name: "p:bg",
                            elements: [
                                {
                                    name: "p:bgPr",
                                    elements: [
                                        {
                                            name: "a:solidFill",
                                            elements: [
                                                {
                                                    name: "a:srgbClr",
                                                    attributes: { val: "FFFFFF" },
                                                },
                                            ],
                                        },
                                    ],
                                },
                            ],
                        },
                        { name: "p:spTree", elements: [] },
                    ],
                },
            ],
        };
        const result = parseSlide(slideXml, mockCtx, "ppt/slides/slide1.xml");
        expect(result.background).toBe("FFFFFF");
    });

    it("should return empty children for missing cSld", () => {
        const result = parseSlide({ elements: [] }, mockCtx, "ppt/slides/slide1.xml");
        expect(result.children).toEqual([]);
    });
});
