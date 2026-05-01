import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathSubScriptHide } from "./math-sub-script-hide";

describe("createMathSubScriptHide", () => {
    describe("#constructor()", () => {
        it("should create a MathSubScriptHide with correct root key", () => {
            const mathSubScriptHide = createMathSubScriptHide();

            const tree = new Formatter().format(mathSubScriptHide);
            expect(tree).to.deep.equal({
                "m:subHide": {
                    _attr: {
                        "m:val": 1,
                    },
                },
            });
        });
    });
});
