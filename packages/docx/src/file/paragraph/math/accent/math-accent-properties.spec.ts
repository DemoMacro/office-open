import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathAccentProperties } from "./math-accent-properties";

describe("createMathAccentProperties", () => {
    it("should create accent properties with accentCharacter", () => {
        const tree = new Formatter().format(
            createMathAccentProperties({ accentCharacter: "\u0302" }),
        );
        expect(tree).to.deep.equal({
            "m:accPr": [{ "m:chr": { _attr: { val: "\u0302" } } }],
        });
    });

    it("should create accent properties with tilde character", () => {
        const tree = new Formatter().format(
            createMathAccentProperties({ accentCharacter: "\u0303" }),
        );
        expect(tree).to.deep.equal({
            "m:accPr": [{ "m:chr": { _attr: { val: "\u0303" } } }],
        });
    });
});
