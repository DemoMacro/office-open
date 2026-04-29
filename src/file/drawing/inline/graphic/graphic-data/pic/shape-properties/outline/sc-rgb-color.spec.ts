import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createScRgbColor } from "./sc-rgb-color";

describe("createScRgbColor", () => {
    it("should create an scRgbColor with basic RGB values", () => {
        const tree = new Formatter().format(createScRgbColor({ r: "100%", g: "0%", b: "0%" }));
        expect(tree).to.deep.equal({
            "a:scrgbClr": { _attr: { r: "100%", g: "0%", b: "0%" } },
        });
    });

    it("should create an scRgbColor with transforms", () => {
        const tree = new Formatter().format(
            createScRgbColor({
                r: "100%",
                g: "0%",
                b: "0%",
                transforms: { alpha: 50000 },
            }),
        );
        expect(tree).to.deep.equal({
            "a:scrgbClr": [
                { _attr: { r: "100%", g: "0%", b: "0%" } },
                {
                    "a:alpha": { _attr: { val: 50000 } },
                },
            ],
        });
    });
});
