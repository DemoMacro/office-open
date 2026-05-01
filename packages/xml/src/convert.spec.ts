import { describe, expect, it } from "vitest";

import { toElement } from "../src/convert";

describe("toElement (convert)", () => {
    it("converts simple text element", () => {
        const result = toElement({ "w:t": "Hello" });
        expect(result).toEqual({
            type: "element",
            name: "w:t",
            elements: [{ type: "text", text: "Hello" }],
        });
    });

    it("converts element with attributes", () => {
        const result = toElement({ "w:pStyle": [{ _attr: { "w:val": "Title" } }] });
        expect(result).toEqual({
            type: "element",
            name: "w:pStyle",
            attributes: { "w:val": "Title" },
        });
    });

    it("converts nested elements", () => {
        const result = toElement({
            "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Hello" }] }],
        });
        expect(result.name).toBe("w:p");
        expect(result.attributes).toEqual({ "w:val": "1" });
        expect(result.elements?.[0]?.name).toBe("w:r");
        expect(result.elements?.[0]?.elements?.[0]?.name).toBe("w:t");
    });

    it("converts null value to empty element", () => {
        const result = toElement({ "w:p": null });
        expect(result).toEqual({ type: "element", name: "w:p" });
    });

    it("converts empty array to element without children", () => {
        const result = toElement({ "w:p": [] });
        expect(result).toEqual({ type: "element", name: "w:p" });
    });

    it("converts numeric values", () => {
        const result = toElement({ "w:val": 42 });
        expect(result.elements?.[0]).toEqual({ type: "text", text: "42" });
    });

    it("converts boolean values", () => {
        const result = toElement({ "w:val": true });
        expect(result.elements?.[0]).toEqual({ type: "text", text: "true" });
    });

    it("converts CDATA", () => {
        const result = toElement({ "w:t": { _attr: {}, _cdata: "data" } });
        expect(result.elements?.[0]).toEqual({ type: "cdata", cdata: "data" });
    });
});
