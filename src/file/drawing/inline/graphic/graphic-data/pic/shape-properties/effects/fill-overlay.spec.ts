import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createFillOverlayEffect } from "./fill-overlay";

describe("createFillOverlayEffect", () => {
    it("should create a fill overlay with solid fill", () => {
        const tree = new Formatter().format(
            createFillOverlayEffect({
                blend: "over",
                solidFill: { value: "FF0000" },
            }),
        );
        expect(tree).to.deep.equal({
            "a:fillOverlay": [
                { _attr: { blend: "over" } },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }] },
            ],
        });
    });

    it("should create a fill overlay with noFill", () => {
        const tree = new Formatter().format(
            createFillOverlayEffect({
                blend: "over",
                noFill: true,
            }),
        );
        expect(tree).to.deep.equal({
            "a:fillOverlay": [{ _attr: { blend: "over" } }, { "a:noFill": {} }],
        });
    });

    it("should create a fill overlay with gradient fill", () => {
        const tree = new Formatter().format(
            createFillOverlayEffect({
                blend: "over",
                gradientFill: {
                    stops: [
                        { color: { value: "000000" }, position: 0 },
                        { color: { value: "FFFFFF" }, position: 100000 },
                    ],
                },
            }),
        );
        expect(tree["a:fillOverlay"]).to.have.length(2);
    });

    it("should create a fill overlay with pattern fill", () => {
        const tree = new Formatter().format(
            createFillOverlayEffect({
                blend: "over",
                patternFill: {
                    pattern: "pct50",
                    foregroundColor: { value: "000000" },
                    backgroundColor: { value: "FFFFFF" },
                },
            }),
        );
        expect(tree["a:fillOverlay"]).to.have.length(2);
    });

    it("should create a fill overlay with group fill", () => {
        const tree = new Formatter().format(
            createFillOverlayEffect({
                blend: "over",
                groupFill: true,
            }),
        );
        expect(tree).to.deep.equal({
            "a:fillOverlay": [{ _attr: { blend: "over" } }, { "a:grpFill": {} }],
        });
    });

    it("should default to solid fill when no fill type specified", () => {
        const tree = new Formatter().format(
            createFillOverlayEffect({
                blend: "over",
            }),
        );
        expect(tree).to.deep.equal({
            "a:fillOverlay": [
                { _attr: { blend: "over" } },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "000000" } } }] },
            ],
        });
    });
});
