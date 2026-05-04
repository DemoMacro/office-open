import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { PptxParseContext } from "./context";
import { parseShape, parseConnector } from "./shape";

const mockCtx: PptxParseContext = {
    zip: new Map(),
    slidePaths: [],
};

const mockRels = new Map<string, { target: string; type: string }>();

describe("parseShape", () => {
    it("should parse shape with name and id", () => {
        const sp: Element = {
            elements: [
                {
                    name: "p:nvSpPr",
                    elements: [{ name: "p:cNvPr", attributes: { id: "2", name: "Shape 1" } }],
                },
                {
                    name: "p:spPr",
                    elements: [
                        {
                            name: "a:xfrm",
                            elements: [
                                { name: "a:off", attributes: { x: "100", y: "200" } },
                                { name: "a:ext", attributes: { cx: "400", cy: "300" } },
                            ],
                        },
                        { name: "a:prstGeom", attributes: { prst: "roundRect" } },
                    ],
                },
                {
                    name: "p:txBody",
                    elements: [
                        { name: "a:bodyPr" },
                        {
                            name: "a:p",
                            elements: [{ name: "a:r", elements: [{ name: "a:t", text: "Hello" }] }],
                        },
                    ],
                },
            ],
        };
        const result = parseShape(sp, mockCtx, mockRels);
        expect(result.$type).toBe("shape");
        expect(result.id).toBe(2);
        expect(result.name).toBe("Shape 1");
        expect(result.x).toBe(100);
        expect(result.y).toBe(200);
        expect(result.geometry).toBe("roundRect");
        expect(result.paragraphs).toHaveLength(1);
    });

    it("should detect placeholder types", () => {
        const sp: Element = {
            elements: [
                {
                    name: "p:nvSpPr",
                    elements: [
                        { name: "p:cNvPr", attributes: { id: "1", name: "Title" } },
                        {
                            name: "p:nvPr",
                            elements: [{ name: "p:ph", attributes: { type: "title" } }],
                        },
                    ],
                },
                { name: "p:spPr", elements: [] },
            ],
        };
        const result = parseShape(sp, mockCtx, mockRels);
        expect(result.placeholder).toBe("title");
    });

    it("should detect body placeholder by idx", () => {
        const sp: Element = {
            elements: [
                {
                    name: "p:nvSpPr",
                    elements: [
                        { name: "p:cNvPr", attributes: { id: "2", name: "Content" } },
                        { name: "p:nvPr", elements: [{ name: "p:ph", attributes: { idx: "1" } }] },
                    ],
                },
                { name: "p:spPr", elements: [] },
            ],
        };
        const result = parseShape(sp, mockCtx, mockRels);
        expect(result.placeholder).toBe("body");
        expect(result.placeholderIndex).toBe(1);
    });
});

describe("parseConnector", () => {
    it("should parse connector shape", () => {
        const cxnSp: Element = {
            elements: [
                {
                    name: "p:nvCxnSpPr",
                    elements: [{ name: "p:cNvPr", attributes: { id: "5", name: "Connector 1" } }],
                },
                {
                    name: "p:spPr",
                    elements: [{ name: "a:prstGeom", attributes: { prst: "bentConnector3" } }],
                },
            ],
        };
        const result = parseConnector(cxnSp, mockCtx);
        expect(result.$type).toBe("connectorShape");
        expect(result.id).toBe(5);
        expect(result.style).toBe("bentConnector3");
    });
});
