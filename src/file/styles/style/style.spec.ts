import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Style } from "./style";

describe("Style", () => {
    describe("#constructor()", () => {
        it("should set the given properties", () => {
            const style = new Style(
                {
                    default: true,
                    styleId: "myStyleId",
                    type: "paragraph",
                },
                {},
            );
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": {
                    _attr: { "w:default": true, "w:styleId": "myStyleId", "w:type": "paragraph" },
                },
            });
        });

        it("should set the name of the style, if given", () => {
            const style = new Style(
                {
                    styleId: "myStyleId",
                    type: "paragraph",
                },
                { name: "Style Name" },
            );
            const tree = new Formatter().format(style);
            expect(tree).to.deep.equal({
                "w:style": [
                    { _attr: { "w:styleId": "myStyleId", "w:type": "paragraph" } },
                    { "w:name": { _attr: { "w:val": "Style Name" } } },
                ],
            });
        });
    });
});
