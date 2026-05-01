import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createBreak } from "./break";

describe("createBreak", () => {
    it("should create a Break element with correct root key", () => {
        const breakElement = createBreak();
        const tree = new Formatter().format(breakElement);
        expect(tree).to.deep.equal({
            "w:br": {},
        });
    });
});
