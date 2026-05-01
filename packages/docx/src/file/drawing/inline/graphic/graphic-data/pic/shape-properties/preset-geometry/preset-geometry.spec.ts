import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { PresetGeometry } from "./preset-geometry";

describe("PresetGeometry", () => {
    it("should create default rectangle geometry", () => {
        const tree = new Formatter().format(new PresetGeometry());
        expect(tree).to.deep.equal({
            "a:prstGeom": [{ _attr: { prst: "rect" } }, { "a:avLst": {} }],
        });
    });

    it("should create geometry with custom preset", () => {
        const tree = new Formatter().format(new PresetGeometry({ preset: "ellipse" }));
        expect(tree).to.deep.equal({
            "a:prstGeom": [{ _attr: { prst: "ellipse" } }, { "a:avLst": {} }],
        });
    });

    it("should create rounded rectangle with adjustment", () => {
        const tree = new Formatter().format(
            new PresetGeometry({
                preset: "roundRect",
                adjustmentValues: [{ name: "adj", formula: "val 16667" }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:prstGeom": [
                { _attr: { prst: "roundRect" } },
                {
                    "a:avLst": [{ "a:gd": { _attr: { name: "adj", fmla: "val 16667" } } }],
                },
            ],
        });
    });

    it("should create geometry with multiple adjustment values", () => {
        const tree = new Formatter().format(
            new PresetGeometry({
                preset: "star5",
                adjustmentValues: [
                    { name: "adj", formula: "val 19098" },
                    { name: "hf", formula: "val 105146" },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:prstGeom": [
                { _attr: { prst: "star5" } },
                {
                    "a:avLst": [
                        { "a:gd": { _attr: { name: "adj", fmla: "val 19098" } } },
                        { "a:gd": { _attr: { name: "hf", fmla: "val 105146" } } },
                    ],
                },
            ],
        });
    });
});
