import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { PageOrientation, createPageSize } from "./page-size";

describe("PageSize", () => {
    describe("#constructor()", () => {
        it("should create page size with portrait", () => {
            const properties = createPageSize({
                height: 200,
                orientation: PageOrientation.PORTRAIT,
                width: 100,
            });
            const tree = new Formatter().format(properties);

            expect(Object.keys(tree)).to.deep.equal(["w:pgSz"]);
            expect(tree["w:pgSz"]).to.deep.equal({
                _attr: { "w:h": 200, "w:orient": "portrait", "w:w": 100 },
            });
        });

        it("should create page size with horizontal and invert the lengths", () => {
            const properties = createPageSize({
                height: 200,
                orientation: PageOrientation.LANDSCAPE,
                width: 100,
            });
            const tree = new Formatter().format(properties);

            expect(Object.keys(tree)).to.deep.equal(["w:pgSz"]);
            expect(tree["w:pgSz"]).to.deep.equal({
                _attr: { "w:h": 100, "w:orient": "landscape", "w:w": 200 },
            });
        });
    });
});
