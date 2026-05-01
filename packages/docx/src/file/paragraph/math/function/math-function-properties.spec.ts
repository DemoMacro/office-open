import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathFunctionProperties } from "./math-function-properties";

describe("MathFunctionProperties", () => {
    describe("#createMathFunctionProperties()", () => {
        it("should create a MathFunctionProperties with correct root key", () => {
            const mathFunctionProperties = createMathFunctionProperties();

            const tree = new Formatter().format(mathFunctionProperties);
            expect(tree).to.deep.equal({
                "m:funcPr": {},
            });
        });
    });
});
