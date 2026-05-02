import { Formatter } from "@export/formatter";
import { TextRun } from "@file/paragraph/run";
import { describe, expect, it } from "vite-plus/test";

import { Bdo, Dir } from "./bidi";

describe("Dir", () => {
    it("should create a dir element with ltr val", () => {
        const dir = new Dir({
            val: "ltr",
            children: [new TextRun("Hello")],
        });
        const tree = new Formatter().format(dir);
        expect(tree).to.deep.equal({
            "w:dir": [
                { _attr: { "w:val": "ltr" } },
                { "w:r": [{ "w:t": [{ _attr: { "xml:space": "preserve" } }, "Hello"] }] },
            ],
        });
    });

    it("should create a dir element with rtl val", () => {
        const dir = new Dir({
            val: "rtl",
            children: [new TextRun("مرحبا")],
        });
        const tree = new Formatter().format(dir);
        expect(tree).to.deep.equal({
            "w:dir": [
                { _attr: { "w:val": "rtl" } },
                { "w:r": [{ "w:t": [{ _attr: { "xml:space": "preserve" } }, "مرحبا"] }] },
            ],
        });
    });

    it("should create a dir element with multiple children", () => {
        const dir = new Dir({
            val: "ltr",
            children: [new TextRun("Hello"), new TextRun(" World")],
        });
        const tree = new Formatter().format(dir);
        expect(tree).to.deep.equal({
            "w:dir": [
                { _attr: { "w:val": "ltr" } },
                { "w:r": [{ "w:t": [{ _attr: { "xml:space": "preserve" } }, "Hello"] }] },
                { "w:r": [{ "w:t": [{ _attr: { "xml:space": "preserve" } }, " World"] }] },
            ],
        });
    });
});

describe("Bdo", () => {
    it("should create a bdo element with ltr val", () => {
        const bdo = new Bdo({
            val: "ltr",
            children: [new TextRun("123")],
        });
        const tree = new Formatter().format(bdo);
        expect(tree).to.deep.equal({
            "w:bdo": [
                { _attr: { "w:val": "ltr" } },
                { "w:r": [{ "w:t": [{ _attr: { "xml:space": "preserve" } }, "123"] }] },
            ],
        });
    });

    it("should create a bdo element with rtl val", () => {
        const bdo = new Bdo({
            val: "rtl",
            children: [new TextRun("ABC")],
        });
        const tree = new Formatter().format(bdo);
        expect(tree).to.deep.equal({
            "w:bdo": [
                { _attr: { "w:val": "rtl" } },
                { "w:r": [{ "w:t": [{ _attr: { "xml:space": "preserve" } }, "ABC"] }] },
            ],
        });
    });
});
