import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathRunProperties } from "./math-run-properties";

describe("createMathRunProperties", () => {
    it("should create empty run properties", () => {
        const tree = new Formatter().format(createMathRunProperties({}));
        expect(tree).to.deep.equal({ "m:rPr": {} });
    });

    it("should create run properties with lit", () => {
        const tree = new Formatter().format(createMathRunProperties({ lit: true }));
        expect(tree).to.deep.equal({
            "m:rPr": [{ "m:lit": {} }],
        });
    });

    it("should create run properties with normal", () => {
        const tree = new Formatter().format(createMathRunProperties({ normal: true }));
        expect(tree).to.deep.equal({
            "m:rPr": [{ "m:nor": {} }],
        });
    });

    it("should create run properties with script", () => {
        const tree = new Formatter().format(createMathRunProperties({ script: "script" }));
        expect(tree).to.deep.equal({
            "m:rPr": [{ "m:scr": { _attr: { val: "script" } } }],
        });
    });

    it("should create run properties with style", () => {
        const tree = new Formatter().format(createMathRunProperties({ style: "bi" }));
        expect(tree).to.deep.equal({
            "m:rPr": [{ "m:sty": { _attr: { val: "bi" } } }],
        });
    });

    it("should create run properties with script and style", () => {
        const tree = new Formatter().format(
            createMathRunProperties({ script: "fraktur", style: "b" }),
        );
        expect(tree).to.deep.equal({
            "m:rPr": [
                { "m:scr": { _attr: { val: "fraktur" } } },
                { "m:sty": { _attr: { val: "b" } } },
            ],
        });
    });

    it("should create run properties with breakAlignment", () => {
        const tree = new Formatter().format(createMathRunProperties({ breakAlignment: 3 }));
        expect(tree).to.deep.equal({
            "m:rPr": [{ "m:brk": { _attr: { "m:alnAt": "3" } } }],
        });
    });

    it("should create run properties with align", () => {
        const tree = new Formatter().format(createMathRunProperties({ align: true }));
        expect(tree).to.deep.equal({
            "m:rPr": [{ "m:aln": {} }],
        });
    });

    it("should create run properties with all options", () => {
        const tree = new Formatter().format(
            createMathRunProperties({
                lit: false,
                normal: true,
                breakAlignment: 1,
                align: false,
            }),
        );
        const children = tree["m:rPr"] as unknown[];
        expect(children).to.have.length(4);
    });
});
