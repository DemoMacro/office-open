import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { AlignmentType, createAlignment } from "./alignment";

describe("Alignment", () => {
    it("should create", () => {
        const alignment = createAlignment(AlignmentType.BOTH);
        const tree = new Formatter().format(alignment);

        expect(tree).to.deep.equal({
            "w:jc": {
                _attr: {
                    "w:val": "both",
                },
            },
        });
    });
});
