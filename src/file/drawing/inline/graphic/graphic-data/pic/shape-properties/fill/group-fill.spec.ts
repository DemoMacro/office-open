import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createGroupFill } from "./group-fill";

describe("createGroupFill", () => {
    it("should create a group fill element", () => {
        const tree = new Formatter().format(createGroupFill());
        expect(tree).to.deep.equal({ "a:grpFill": {} });
    });
});
