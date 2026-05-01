import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathBoxProperties } from "./math-box-properties";

describe("createMathBoxProperties", () => {
    it("should create empty box properties", () => {
        const tree = new Formatter().format(createMathBoxProperties({}));
        expect(tree).to.deep.equal({ "m:boxPr": {} });
    });

    it("should create box properties with opEmu", () => {
        const tree = new Formatter().format(createMathBoxProperties({ opEmu: true }));
        expect(tree).to.deep.equal({
            "m:boxPr": [{ "m:opEmu": {} }],
        });
    });

    it("should create box properties with noBreak", () => {
        const tree = new Formatter().format(createMathBoxProperties({ noBreak: true }));
        expect(tree).to.deep.equal({
            "m:boxPr": [{ "m:noBreak": {} }],
        });
    });

    it("should create box properties with diff", () => {
        const tree = new Formatter().format(createMathBoxProperties({ diff: false }));
        expect(tree).to.deep.equal({
            "m:boxPr": [{ "m:diff": { _attr: { "w:val": false } } }],
        });
    });

    it("should create box properties with aln", () => {
        const tree = new Formatter().format(createMathBoxProperties({ aln: true }));
        expect(tree).to.deep.equal({
            "m:boxPr": [{ "m:aln": {} }],
        });
    });

    it("should create box properties with all options", () => {
        const tree = new Formatter().format(
            createMathBoxProperties({ opEmu: true, noBreak: false, diff: true, aln: false }),
        );
        const children = tree["m:boxPr"] as unknown[];
        expect(children).to.have.length(4);
    });
});
