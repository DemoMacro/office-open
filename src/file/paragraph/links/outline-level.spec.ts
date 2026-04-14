import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createOutlineLevel } from "./outline-level";

describe("ParagraphOutlineLevel", () => {
    describe("#createOutlineLevel()", () => {
        it("should create an outlineLevel with given value", () => {
            const outlineLevel = createOutlineLevel(0);
            const tree = new Formatter().format(outlineLevel);
            expect(tree).to.deep.equal({
                "w:outlineLvl": {
                    _attr: {
                        "w:val": 0,
                    },
                },
            });
        });
    });
});
