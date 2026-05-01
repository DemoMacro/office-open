import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createCnfStyle } from "./cnf-style";

describe("createCnfStyle", () => {
    it("should create a cnfStyle with no attributes when empty options", () => {
        const tree = new Formatter().format(createCnfStyle({}));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: {} } });
    });

    it("should create a cnfStyle with firstRow=true", () => {
        const tree = new Formatter().format(createCnfStyle({ firstRow: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:firstRow": "1" } } });
    });

    it("should create a cnfStyle with firstRow=false", () => {
        const tree = new Formatter().format(createCnfStyle({ firstRow: false }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:firstRow": "0" } } });
    });

    it("should create a cnfStyle with lastRow", () => {
        const tree = new Formatter().format(createCnfStyle({ lastRow: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:lastRow": "1" } } });
    });

    it("should create a cnfStyle with firstColumn", () => {
        const tree = new Formatter().format(createCnfStyle({ firstColumn: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:firstColumn": "1" } } });
    });

    it("should create a cnfStyle with lastColumn", () => {
        const tree = new Formatter().format(createCnfStyle({ lastColumn: false }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:lastColumn": "0" } } });
    });

    it("should create a cnfStyle with oddVBand", () => {
        const tree = new Formatter().format(createCnfStyle({ oddVBand: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:oddVBand": "1" } } });
    });

    it("should create a cnfStyle with evenVBand", () => {
        const tree = new Formatter().format(createCnfStyle({ evenVBand: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:evenVBand": "1" } } });
    });

    it("should create a cnfStyle with oddHBand", () => {
        const tree = new Formatter().format(createCnfStyle({ oddHBand: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:oddHBand": "1" } } });
    });

    it("should create a cnfStyle with evenHBand", () => {
        const tree = new Formatter().format(createCnfStyle({ evenHBand: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:evenHBand": "1" } } });
    });

    it("should create a cnfStyle with firstRowFirstColumn", () => {
        const tree = new Formatter().format(createCnfStyle({ firstRowFirstColumn: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:firstRowFirstColumn": "1" } } });
    });

    it("should create a cnfStyle with firstRowLastColumn", () => {
        const tree = new Formatter().format(createCnfStyle({ firstRowLastColumn: false }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:firstRowLastColumn": "0" } } });
    });

    it("should create a cnfStyle with lastRowFirstColumn", () => {
        const tree = new Formatter().format(createCnfStyle({ lastRowFirstColumn: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:lastRowFirstColumn": "1" } } });
    });

    it("should create a cnfStyle with lastRowLastColumn", () => {
        const tree = new Formatter().format(createCnfStyle({ lastRowLastColumn: true }));
        expect(tree).to.deep.equal({ "w:cnfStyle": { _attr: { "w:lastRowLastColumn": "1" } } });
    });

    it("should create a cnfStyle with multiple attributes", () => {
        const tree = new Formatter().format(
            createCnfStyle({
                firstRow: true,
                firstColumn: true,
                firstRowFirstColumn: true,
            }),
        );
        expect(tree).to.deep.equal({
            "w:cnfStyle": {
                _attr: {
                    "w:firstRow": "1",
                    "w:firstColumn": "1",
                    "w:firstRowFirstColumn": "1",
                },
            },
        });
    });
});
