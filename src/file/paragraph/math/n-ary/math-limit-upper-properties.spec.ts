import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathLimitUpperProperties } from "./math-limit-upper-properties";

describe("createMathLimitUpperProperties", () => {
    it("should create empty limit upper properties", () => {
        const tree = new Formatter().format(createMathLimitUpperProperties());
        expect(tree).to.deep.equal({ "m:limUppPr": {} });
    });

    it("should create limit upper properties with undefined options", () => {
        const tree = new Formatter().format(createMathLimitUpperProperties(undefined));
        expect(tree).to.deep.equal({ "m:limUppPr": {} });
    });
});
