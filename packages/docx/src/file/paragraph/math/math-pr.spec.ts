import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathProperties } from "./math-pr";

describe("createMathProperties", () => {
    it("should create empty math properties", () => {
        const tree = new Formatter().format(createMathProperties({}));
        expect(tree).to.deep.equal({ "m:mathPr": {} });
    });

    it("should create math properties with mathFont", () => {
        const tree = new Formatter().format(createMathProperties({ mathFont: "Cambria Math" }));
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:mathFont": { _attr: { "m:val": "Cambria Math" } } }],
        });
    });

    it("should create math properties with breakBin", () => {
        const tree = new Formatter().format(createMathProperties({ breakBin: "before" }));
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:brkBin": { _attr: { "m:val": "before" } } }],
        });
    });

    it("should create math properties with breakBinSub", () => {
        const tree = new Formatter().format(createMathProperties({ breakBinSub: "--" }));
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:brkBinSub": { _attr: { "m:val": "--" } } }],
        });
    });

    it("should create math properties with smallFrac", () => {
        const tree = new Formatter().format(createMathProperties({ smallFrac: true }));
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:smallFrac": {} }],
        });
    });

    it("should create math properties with displayDefault", () => {
        const tree = new Formatter().format(createMathProperties({ displayDefault: true }));
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:dispDef": {} }],
        });
    });

    it("should create math properties with leftMargin and rightMargin", () => {
        const tree = new Formatter().format(
            createMathProperties({ leftMargin: 100, rightMargin: 200 }),
        );
        expect(tree).to.deep.equal({
            "m:mathPr": [
                { "m:lMargin": { _attr: { "m:val": "100" } } },
                { "m:rMargin": { _attr: { "m:val": "200" } } },
            ],
        });
    });

    it("should create math properties with defaultJustification", () => {
        const tree = new Formatter().format(
            createMathProperties({ defaultJustification: "center" }),
        );
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:defJc": { _attr: { "m:val": "center" } } }],
        });
    });

    it("should create math properties with spacing options", () => {
        const tree = new Formatter().format(
            createMathProperties({
                preSpacing: 50,
                postSpacing: 50,
                interSpacing: 100,
                intraSpacing: 25,
            }),
        );
        expect(tree).to.deep.equal({
            "m:mathPr": [
                { "m:preSp": { _attr: { "m:val": "50" } } },
                { "m:postSp": { _attr: { "m:val": "50" } } },
                { "m:interSp": { _attr: { "m:val": "100" } } },
                { "m:intraSp": { _attr: { "m:val": "25" } } },
            ],
        });
    });

    it("should create math properties with wrapIndent and wrapRight", () => {
        const tree = new Formatter().format(
            createMathProperties({ wrapIndent: 500, wrapRight: true }),
        );
        expect(tree).to.deep.equal({
            "m:mathPr": [{ "m:wrapIndent": { _attr: { "m:val": "500" } } }, { "m:wrapRight": {} }],
        });
    });

    it("should create math properties with limit locations", () => {
        const tree = new Formatter().format(
            createMathProperties({
                integralLimitLocation: "subSup",
                naryLimitLocation: "undOvr",
            }),
        );
        expect(tree).to.deep.equal({
            "m:mathPr": [
                { "m:intLim": { _attr: { "m:val": "subSup" } } },
                { "m:naryLim": { _attr: { "m:val": "undOvr" } } },
            ],
        });
    });

    it("should create math properties with all options", () => {
        const tree = new Formatter().format(
            createMathProperties({
                mathFont: "Cambria Math",
                breakBin: "before",
                breakBinSub: "--",
                smallFrac: false,
                displayDefault: true,
                leftMargin: 0,
                rightMargin: 0,
                defaultJustification: "center",
                preSpacing: 0,
                postSpacing: 0,
                interSpacing: 0,
                intraSpacing: 0,
                wrapIndent: 0,
                wrapRight: false,
                integralLimitLocation: "subSup",
                naryLimitLocation: "undOvr",
            }),
        );
        const children = tree["m:mathPr"] as unknown[];
        expect(children).to.have.length(16);
    });
});
