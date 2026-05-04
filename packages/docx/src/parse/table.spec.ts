import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxParseContext } from "./context";
import { parseTable } from "./table";

const mockCtx: DocxParseContext = {
    zip: new Map(),
    hyperlinks: new Map(),
    media: new Map(),
    documentRels: new Map(),
};

describe("parseTable", () => {
    it("should parse a simple table", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:p",
                                    elements: [
                                        {
                                            name: "w:r",
                                            elements: [{ name: "w:t", text: "Cell 1" }],
                                        },
                                    ],
                                },
                            ],
                        },
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:p",
                                    elements: [
                                        {
                                            name: "w:r",
                                            elements: [{ name: "w:t", text: "Cell 2" }],
                                        },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.$type).toBe("table");
        expect(result.rows).toHaveLength(1);
        expect(result.rows[0].cells).toHaveLength(2);
        expect(result.rows[0].cells[0].children![0].text).toBe("Cell 1");
        expect(result.rows[0].cells[1].children![0].text).toBe("Cell 2");
    });

    it("should parse table width", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tblPr",
                    elements: [{ name: "w:tblW", attributes: { "w:w": "5000", "w:type": "pct" } }],
                },
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:p",
                                    elements: [
                                        { name: "w:r", elements: [{ name: "w:t", text: "A" }] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.width).toEqual({ size: 5000, type: "pct" });
    });

    it("should parse table style", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tblPr",
                    elements: [{ name: "w:tblStyle", attributes: { "w:val": "TableGrid" } }],
                },
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:p",
                                    elements: [
                                        { name: "w:r", elements: [{ name: "w:t", text: "A" }] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.style).toBe("TableGrid");
    });

    it("should parse column span", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:tcPr",
                                    elements: [
                                        { name: "w:gridSpan", attributes: { "w:val": "2" } },
                                    ],
                                },
                                {
                                    name: "w:p",
                                    elements: [
                                        { name: "w:r", elements: [{ name: "w:t", text: "Span" }] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.rows[0].cells[0].columnSpan).toBe(2);
    });

    it("should parse cell shading", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:tcPr",
                                    elements: [
                                        { name: "w:shd", attributes: { "w:fill": "DDDDDD" } },
                                    ],
                                },
                                {
                                    name: "w:p",
                                    elements: [
                                        { name: "w:r", elements: [{ name: "w:t", text: "Gray" }] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.rows[0].cells[0].shading).toEqual({ fill: "DDDDDD" });
    });

    it("should parse cell vertical alignment", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:tcPr",
                                    elements: [
                                        { name: "w:vAlign", attributes: { "w:val": "center" } },
                                    ],
                                },
                                {
                                    name: "w:p",
                                    elements: [
                                        { name: "w:r", elements: [{ name: "w:t", text: "V" }] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.rows[0].cells[0].verticalAlign).toBe("center");
    });

    it("should parse cell width", () => {
        const tbl: Element = {
            elements: [
                {
                    name: "w:tr",
                    elements: [
                        {
                            name: "w:tc",
                            elements: [
                                {
                                    name: "w:tcPr",
                                    elements: [
                                        {
                                            name: "w:tcW",
                                            attributes: { "w:w": "2000", "w:type": "dxa" },
                                        },
                                    ],
                                },
                                {
                                    name: "w:p",
                                    elements: [
                                        { name: "w:r", elements: [{ name: "w:t", text: "W" }] },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseTable(tbl, mockCtx);
        expect(result.rows[0].cells[0].width).toEqual({ size: 2000, type: "dxa" });
    });
});
