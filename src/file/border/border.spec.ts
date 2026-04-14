import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { describe, expect, it } from "vite-plus/test";

import { createBorderElement } from "./border";

describe("BorderElement", () => {
    describe("#createBorderElement", () => {
        it("should create a simple border element", () => {
            const border = createBorderElement("w:top", {
                style: BorderStyle.SINGLE,
            });
            const tree = new Formatter().format(border);
            expect(tree).to.deep.equal({
                "w:top": {
                    _attr: {
                        "w:val": "single",
                    },
                },
            });
        });
        it("should create a simple border element with a size", () => {
            const border = createBorderElement("w:top", {
                size: 22,
                style: BorderStyle.SINGLE,
            });
            const tree = new Formatter().format(border);
            expect(tree).to.deep.equal({
                "w:top": {
                    _attr: {
                        "w:sz": 22,
                        "w:val": "single",
                    },
                },
            });
        });
        it("should create a simple border element with space", () => {
            const border = createBorderElement("w:top", {
                space: 22,
                style: BorderStyle.SINGLE,
            });
            const tree = new Formatter().format(border);
            expect(tree).to.deep.equal({
                "w:top": {
                    _attr: {
                        "w:space": 22,
                        "w:val": "single",
                    },
                },
            });
        });
    });
});
