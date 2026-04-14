import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createLineNumberType } from "./line-number";

describe("createLineNumberType", () => {
    it("should work", () => {
        const textDirection = createLineNumberType({
            countBy: 0,
            distance: 10,
            restart: "newPage",
            start: 0,
        });

        const tree = new Formatter().format(textDirection);
        expect(tree).to.deep.equal({
            "w:lnNumType": {
                _attr: { "w:countBy": 0, "w:distance": 10, "w:restart": "newPage", "w:start": 0 },
            },
        });
    });

    it("should work with string measures for distance", () => {
        const textDirection = createLineNumberType({
            countBy: 0,
            distance: "10mm",
            restart: "newPage",
            start: 0,
        });

        const tree = new Formatter().format(textDirection);
        expect(tree).to.deep.equal({
            "w:lnNumType": {
                _attr: {
                    "w:countBy": 0,
                    "w:distance": "10mm",
                    "w:restart": "newPage",
                    "w:start": 0,
                },
            },
        });
    });

    it("should work with blank entries", () => {
        const textDirection = createLineNumberType({});

        const tree = new Formatter().format(textDirection);
        expect(tree).to.deep.equal({
            "w:lnNumType": { _attr: {} },
        });
    });
});
