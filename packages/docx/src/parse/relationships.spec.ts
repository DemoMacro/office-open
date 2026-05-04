import { parseRels, findRel, findRelsByType } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

describe("parseRels", () => {
    it("should return empty array for missing file", () => {
        const result = parseRels(new Map(), "missing.xml");
        expect(result).toEqual([]);
    });

    it("should return empty array for nonexistent rels file", () => {
        const zip = new Map([["word/_rels/document.xml.rels", new Uint8Array()]]);
        const result = parseRels(zip, "nonexistent.rels");
        expect(result).toEqual([]);
    });
});

describe("findRel", () => {
    it("should find relationship by id", () => {
        const rels = [
            { id: "rId1", target: "image.png", type: "http://example.com/image" },
            { id: "rId2", target: "style.css", type: "http://example.com/style" },
        ];
        expect(findRel(rels, "rId1")?.target).toBe("image.png");
    });

    it("should return undefined for missing id", () => {
        expect(findRel([], "rId99")).toBeUndefined();
    });
});

describe("findRelsByType", () => {
    it("should filter by type substring", () => {
        const rels = [
            { id: "rId1", target: "a.png", type: "http://example.com/image" },
            { id: "rId2", target: "https://x.com", type: "http://example.com/hyperlink" },
        ];
        expect(findRelsByType(rels, "hyperlink")).toHaveLength(1);
        expect(findRelsByType(rels, "image")).toHaveLength(1);
    });
});
