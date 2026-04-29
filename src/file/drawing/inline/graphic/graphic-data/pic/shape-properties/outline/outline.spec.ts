import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createOutline } from "./outline";
import { SchemeColor } from "./scheme-color";

describe("createOutline", () => {
    it("should create no fill", () => {
        const tree = new Formatter().format(createOutline({ type: "noFill" }));
        expect(tree).to.deep.equal({
            "a:ln": [
                {
                    _attr: {},
                },
                {
                    "a:noFill": {},
                },
            ],
        });
    });

    it("should create solid rgb fill", () => {
        const tree = new Formatter().format(
            createOutline({ type: "solidFill", color: { value: "FFFFFF" } }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                {
                    _attr: {},
                },
                {
                    "a:solidFill": [
                        {
                            "a:srgbClr": {
                                _attr: {
                                    val: "FFFFFF",
                                },
                            },
                        },
                    ],
                },
            ],
        });
    });

    it("should create solid scheme fill", () => {
        const tree = new Formatter().format(
            createOutline({
                type: "solidFill",
                color: { value: SchemeColor.ACCENT1 },
            }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                {
                    _attr: {},
                },
                {
                    "a:solidFill": [
                        {
                            "a:schemeClr": {
                                _attr: {
                                    val: "accent1",
                                },
                            },
                        },
                    ],
                },
            ],
        });
    });

    it("should create outline with dash", () => {
        const tree = new Formatter().format(
            createOutline({ type: "solidFill", color: { value: "000000" }, dash: "DASH" }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: {} },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "000000" } } }] },
                { "a:prstDash": { _attr: { val: "dash" } } },
            ],
        });
    });

    it("should create outline with round join", () => {
        const tree = new Formatter().format(
            createOutline({ type: "solidFill", color: { value: "000000" }, join: "ROUND" }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: {} },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "000000" } } }] },
                { "a:round": {} },
            ],
        });
    });

    it("should create outline with miter join and limit", () => {
        const tree = new Formatter().format(
            createOutline({
                type: "solidFill",
                color: { value: "000000" },
                join: "MITER",
                miterLimit: 800000,
            }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: {} },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "000000" } } }] },
                { "a:miter": { _attr: { lim: 800000 } } },
            ],
        });
    });

    it("should create outline with width, cap, compoundLine, and dash", () => {
        const tree = new Formatter().format(
            createOutline({
                type: "solidFill",
                color: { value: "FF0000" },
                width: 19050,
                cap: "ROUND",
                compoundLine: "DOUBLE",
                dash: "DOT",
                align: "INSET",
            }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: { algn: "in", cap: "rnd", cmpd: "dbl", w: 19050 } },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "FF0000" } } }] },
                { "a:prstDash": { _attr: { val: "dot" } } },
            ],
        });
    });

    it("should create outline with width only (no fill)", () => {
        const tree = new Formatter().format(createOutline({ width: 12700 }));
        expect(tree).to.deep.equal({
            "a:ln": { _attr: { w: 12700 } },
        });
    });

    it("should create outline with custom dash pattern", () => {
        const tree = new Formatter().format(
            createOutline({
                type: "solidFill",
                color: { value: "000000" },
                customDash: [{ d: "500%", sp: "200%" }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: {} },
                { "a:solidFill": [{ "a:srgbClr": { _attr: { val: "000000" } } }] },
                {
                    "a:custDash": [
                        {
                            "a:ds": { _attr: { d: "500%", sp: "200%" } },
                        },
                    ],
                },
            ],
        });
    });

    it("should create outline with multiple custom dash stops", () => {
        const tree = new Formatter().format(
            createOutline({
                customDash: [
                    { d: "500%", sp: "200%" },
                    { d: "100%", sp: "200%" },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: {} },
                {
                    "a:custDash": [
                        { "a:ds": { _attr: { d: "500%", sp: "200%" } } },
                        { "a:ds": { _attr: { d: "100%", sp: "200%" } } },
                    ],
                },
            ],
        });
    });

    it("should prefer customDash over dash when both specified", () => {
        const tree = new Formatter().format(
            createOutline({
                customDash: [{ d: "300%", sp: "100%" }],
                dash: "DASH",
            }),
        );
        expect(tree).to.deep.equal({
            "a:ln": [
                { _attr: {} },
                {
                    "a:custDash": [{ "a:ds": { _attr: { d: "300%", sp: "100%" } } }],
                },
            ],
        });
    });
});
