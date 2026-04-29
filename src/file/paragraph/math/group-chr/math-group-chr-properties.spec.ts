import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathGroupChrProperties } from "./math-group-chr-properties";

describe("createMathGroupChrProperties", () => {
    it("should create empty group chr properties", () => {
        const tree = new Formatter().format(createMathGroupChrProperties({}));
        expect(tree).to.deep.equal({ "m:groupChrPr": {} });
    });

    it("should create group chr properties with chr", () => {
        const tree = new Formatter().format(createMathGroupChrProperties({ chr: "\u007B" }));
        expect(tree).to.deep.equal({
            "m:groupChrPr": [{ "m:chr": { _attr: { "m:val": "\u007B" } } }],
        });
    });

    it("should create group chr properties with pos", () => {
        const tree = new Formatter().format(createMathGroupChrProperties({ pos: "bot" }));
        expect(tree).to.deep.equal({
            "m:groupChrPr": [{ "m:pos": { _attr: { "m:val": "bot" } } }],
        });
    });

    it("should create group chr properties with vertJc", () => {
        const tree = new Formatter().format(createMathGroupChrProperties({ vertJc: "top" }));
        expect(tree).to.deep.equal({
            "m:groupChrPr": [{ "m:vertJc": { _attr: { "m:val": "top" } } }],
        });
    });

    it("should create group chr properties with all options", () => {
        const tree = new Formatter().format(
            createMathGroupChrProperties({ chr: "\u23DF", pos: "top", vertJc: "bot" }),
        );
        expect(tree).to.deep.equal({
            "m:groupChrPr": [
                { "m:chr": { _attr: { "m:val": "\u23DF" } } },
                { "m:pos": { _attr: { "m:val": "top" } } },
                { "m:vertJc": { _attr: { "m:val": "bot" } } },
            ],
        });
    });
});
