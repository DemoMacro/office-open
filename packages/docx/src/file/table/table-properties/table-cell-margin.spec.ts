import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { WidthType } from "../table-width";
import { createCellMargin, createTableCellMargin } from "./table-cell-margin";

describe("TableCellMargin", () => {
    describe("#createTableCellMargin", () => {
        it("should return undefined if there are no margins specified", () => {
            const cellMargin = createTableCellMargin({});
            expect(cellMargin).to.be.undefined;
        });

        it("should add a table cell top margin", () => {
            const cellMargin = createTableCellMargin({
                marginUnitType: WidthType.DXA,
                top: 1234,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:top": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell top margin using default width type", () => {
            const cellMargin = createTableCellMargin({
                top: 1234,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:top": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell left margin", () => {
            const cellMargin = createTableCellMargin({
                left: 1234,
                marginUnitType: WidthType.DXA,
            });
            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:left": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell left margin using default width type", () => {
            const cellMargin = createTableCellMargin({
                left: 1234,
            });
            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:left": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell bottom margin", () => {
            const cellMargin = createTableCellMargin({
                bottom: 1234,
                marginUnitType: WidthType.DXA,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:bottom": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell bottom margin using default width type", () => {
            const cellMargin = createTableCellMargin({
                bottom: 1234,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:bottom": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell right margin", () => {
            const cellMargin = createTableCellMargin({
                marginUnitType: WidthType.DXA,
                right: 1234,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:right": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a table cell right margin using default width type", () => {
            const cellMargin = createTableCellMargin({
                right: 1234,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tblCellMar": [{ "w:right": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });
    });

    describe("#createCellMargin", () => {
        it("should return undefined if there are no margins specified", () => {
            const cellMargin = createCellMargin({});
            expect(cellMargin).to.be.undefined;
        });

        it("should create cell-level margin element with tcMar tag", () => {
            const cellMargin = createCellMargin({
                bottom: 100,
                top: 100,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tcMar": [
                    { "w:top": { _attr: { "w:type": "dxa", "w:w": 100 } } },
                    { "w:bottom": { _attr: { "w:type": "dxa", "w:w": 100 } } },
                ],
            });
        });

        it("should support all margin positions", () => {
            const cellMargin = createCellMargin({
                bottom: 50,
                left: 100,
                right: 100,
                top: 50,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tcMar": [
                    { "w:top": { _attr: { "w:type": "dxa", "w:w": 50 } } },
                    { "w:left": { _attr: { "w:type": "dxa", "w:w": 100 } } },
                    { "w:bottom": { _attr: { "w:type": "dxa", "w:w": 50 } } },
                    { "w:right": { _attr: { "w:type": "dxa", "w:w": 100 } } },
                ],
            });
        });

        it("should support custom width type", () => {
            const cellMargin = createCellMargin({
                left: 5,
                marginUnitType: WidthType.PERCENTAGE,
            });

            const tree = new Formatter().format(cellMargin!);
            expect(tree).to.deep.equal({
                "w:tcMar": [{ "w:left": { _attr: { "w:type": "pct", "w:w": "5%" } } }],
            });
        });
    });
});
