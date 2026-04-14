import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathEndingCharacter } from "./math-ending-char";

describe("createMathEndingCharacter", () => {
    describe("#constructor()", () => {
        it("should create a MathEndingCharacter with correct root key", () => {
            const mathEndingCharacter = createMathEndingCharacter({ character: "]" });

            const tree = new Formatter().format(mathEndingCharacter);
            expect(tree).to.deep.equal({
                "m:endChr": {
                    _attr: {
                        "m:val": "]",
                    },
                },
            });
        });
    });
});
