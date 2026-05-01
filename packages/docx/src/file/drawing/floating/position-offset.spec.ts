import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createPositionOffset } from "./position-offset";

describe("createPositionOffset", () => {
    describe("#constructor()", () => {
        it("should create a element with correct root key", () => {
            const tree = new Formatter().format(createPositionOffset(50));
            expect(tree).to.deep.equal({
                "wp:posOffset": ["50"],
            });
        });
    });
});
