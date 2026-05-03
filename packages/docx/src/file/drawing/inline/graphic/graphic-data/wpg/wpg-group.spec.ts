import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createWpgGroup } from "./wpg-group";

const TRANSFORM = {
    emus: { x: 914400, y: 914400 },
    pixels: { x: 100, y: 100 },
};

describe("WpgGroup", () => {
    it("should create group with chOff and chExt", () => {
        const tree = new Formatter().format(
            createWpgGroup({
                children: [],
                transformation: TRANSFORM,
                chOff: { x: 100000, y: 50000 },
                chExt: { cx: 800000, cy: 600000 },
            }),
        );
        const grpSpPr = tree["wpg:wgp"][1]["wpg:grpSpPr"];
        expect(grpSpPr).to.include.deep.members([
            {
                "a:xfrm": [
                    {
                        "a:off": {
                            _attr: { x: 0, y: 0 },
                        },
                    },
                    {
                        "a:ext": {
                            _attr: { cx: 914400, cy: 914400 },
                        },
                    },
                    {
                        "a:chOff": {
                            _attr: { x: 100000, y: 50000 },
                        },
                    },
                    {
                        "a:chExt": {
                            _attr: { cx: 800000, cy: 600000 },
                        },
                    },
                ],
            },
        ]);
    });

    it("should create group with solid fill", () => {
        const tree = new Formatter().format(
            createWpgGroup({
                children: [],
                transformation: TRANSFORM,
                fill: "FF0000",
            }),
        );
        const grpSpPr = tree["wpg:wgp"][1]["wpg:grpSpPr"];
        expect(grpSpPr).to.include.deep.members([
            { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }] },
        ]);
    });

    it("should create group with no fill", () => {
        const tree = new Formatter().format(
            createWpgGroup({
                children: [],
                transformation: TRANSFORM,
                fill: { type: "none" },
            }),
        );
        const grpSpPr = tree["wpg:wgp"][1]["wpg:grpSpPr"];
        expect(grpSpPr).to.include.deep.members([{ "a:noFill": {} }]);
    });

    it("should create group with effects", () => {
        const tree = new Formatter().format(
            createWpgGroup({
                children: [],
                transformation: TRANSFORM,
                effects: {
                    softEdge: 50800,
                },
            }),
        );
        const grpSpPr = tree["wpg:wgp"][1]["wpg:grpSpPr"];
        expect(grpSpPr).to.include.deep.members([
            { "a:effectLst": [{ "a:softEdge": { _attr: { rad: 50800 } } }] },
        ]);
    });
});
