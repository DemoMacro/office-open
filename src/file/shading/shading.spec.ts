import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { ShadingType, createShading } from "./shading";

describe("Shading", () => {
    describe("#createShading", () => {
        it("should create", () => {
            const shading = createShading({});
            const tree = new Formatter().format(shading);
            expect(tree).to.deep.equal({
                "w:shd": {
                    _attr: {},
                },
            });
        });

        it("should create with params", () => {
            const shading = createShading({
                color: "FF0000",
                fill: "555555",
                type: ShadingType.PERCENT_40,
            });
            const tree = new Formatter().format(shading);
            expect(tree).to.deep.equal({
                "w:shd": {
                    _attr: {
                        "w:color": "FF0000",
                        "w:fill": "555555",
                        "w:val": "pct40",
                    },
                },
            });
        });
    });
});
