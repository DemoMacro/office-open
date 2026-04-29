import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createDivId } from "./div-id";

describe("createDivId", () => {
    it("should create a divId element with the given id", () => {
        const tree = new Formatter().format(createDivId(1));
        expect(tree).to.deep.equal({ "w:divId": { _attr: { "w:val": 1 } } });
    });

    it("should create a divId element with id 0", () => {
        const tree = new Formatter().format(createDivId(0));
        expect(tree).to.deep.equal({ "w:divId": { _attr: { "w:val": 0 } } });
    });
});
