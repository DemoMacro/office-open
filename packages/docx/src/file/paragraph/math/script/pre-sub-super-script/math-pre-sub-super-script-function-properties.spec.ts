import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathPreSubSuperScriptProperties } from "./math-pre-sub-super-script-function-properties";

describe("createMathPreSubSuperScriptProperties", () => {
    describe("#constructor()", () => {
        it("should create a MathPreSubSuperScriptProperties with correct root key", () => {
            const mathPreSubSuperScriptProperties = createMathPreSubSuperScriptProperties();

            const tree = new Formatter().format(mathPreSubSuperScriptProperties);
            expect(tree).to.deep.equal({
                "m:sPrePr": {},
            });
        });
    });
});
