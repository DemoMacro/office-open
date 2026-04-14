import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathSuperScriptHide } from "./math-super-script-hide";

describe("createMathSuperScriptHide", () => {
    describe("#constructor()", () => {
        it("should create a MathSuperScriptHide with correct root key", () => {
            const mathSuperScriptHide = createMathSuperScriptHide();

            const tree = new Formatter().format(mathSuperScriptHide);
            expect(tree).to.deep.equal({
                "m:supHide": {
                    _attr: {
                        "m:val": 1,
                    },
                },
            });
        });
    });
});
