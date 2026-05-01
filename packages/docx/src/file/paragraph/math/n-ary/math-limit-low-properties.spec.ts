import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathLimitLowProperties } from "./math-limit-low-properties";

describe("createMathLimitLowProperties", () => {
    it("should create empty limit low properties", () => {
        const tree = new Formatter().format(createMathLimitLowProperties());
        expect(tree).to.deep.equal({ "m:limLowPr": {} });
    });

    it("should create limit low properties with undefined options", () => {
        const tree = new Formatter().format(createMathLimitLowProperties(undefined));
        expect(tree).to.deep.equal({ "m:limLowPr": {} });
    });
});
