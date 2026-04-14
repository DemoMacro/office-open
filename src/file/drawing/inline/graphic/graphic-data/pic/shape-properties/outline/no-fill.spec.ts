import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createNoFill } from "./no-fill";

describe("NoFill", () => {
    describe("#constructor()", () => {
        it("should create", () => {
            const tree = new Formatter().format(createNoFill());
            expect(tree).to.deep.equal({
                "a:noFill": {},
            });
        });
    });
});
