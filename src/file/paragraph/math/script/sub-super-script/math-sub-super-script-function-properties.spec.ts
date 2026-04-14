import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathSubSuperScriptProperties } from "./math-sub-super-script-function-properties";

describe("createMathSuperScriptProperties", () => {
    describe("#constructor()", () => {
        it("should create a MathSubSuperScriptProperties with correct root key", () => {
            const mathSubSuperScriptProperties = createMathSubSuperScriptProperties();

            const tree = new Formatter().format(mathSubSuperScriptProperties);
            expect(tree).to.deep.equal({
                "m:sSubSupPr": {},
            });
        });
    });
});
