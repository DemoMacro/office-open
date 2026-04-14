import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { VerticalAlignTable } from "@file/vertical-align";
import { describe, expect, it } from "vite-plus/test";

import { WidthType } from "../table-width";
import { VerticalMergeType } from "./table-cell-components";
import { TableCellProperties } from "./table-cell-properties";

describe("TableCellProperties", () => {
    describe("#constructor", () => {
        it("creates an initially empty property object", () => {
            const properties = new TableCellProperties({});
            // The TableCellProperties is ignorable if there are no attributes,
            // Which results in prepForXml returning undefined, which causes
            // The formatter to throw an error if that is the only object it
            // Has been asked to format.
            expect(() => new Formatter().format(properties)).to.throw(
                "XMLComponent did not format correctly",
            );
        });

        it("adds grid span", () => {
            const properties = new TableCellProperties({ columnSpan: 1 });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({ "w:tcPr": [{ "w:gridSpan": { _attr: { "w:val": 1 } } }] });
        });

        it("adds vertical merge", () => {
            const properties = new TableCellProperties({
                verticalMerge: VerticalMergeType.CONTINUE,
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:tcPr": [{ "w:vMerge": { _attr: { "w:val": "continue" } } }],
            });
        });

        it("sets vertical align", () => {
            const properties = new TableCellProperties({
                verticalAlign: VerticalAlignTable.BOTTOM,
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:tcPr": [{ "w:vAlign": { _attr: { "w:val": "bottom" } } }],
            });
        });

        it("should set width", () => {
            const properties = new TableCellProperties({
                width: {
                    size: 1,
                    type: WidthType.DXA,
                },
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:tcPr": [{ "w:tcW": { _attr: { "w:type": "dxa", "w:w": 1 } } }],
            });
        });

        it("should set width using default of AUTO", () => {
            const properties = new TableCellProperties({
                width: {
                    size: 1,
                },
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:tcPr": [{ "w:tcW": { _attr: { "w:type": "auto", "w:w": 1 } } }],
            });
        });

        it("sets shading", () => {
            const properties = new TableCellProperties({
                shading: {
                    color: "000000",
                    fill: "ffffff",
                },
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal({
                "w:tcPr": [{ "w:shd": { _attr: { "w:color": "000000", "w:fill": "ffffff" } } }],
            });
        });

        it("should set the TableCellBorders", () => {
            const properties = new TableCellProperties({
                borders: {
                    top: {
                        color: "ff0000",
                        size: 3,
                        style: BorderStyle.DASH_DOT_STROKED,
                    },
                },
            });

            const tree = new Formatter().format(properties);

            expect(tree["w:tcPr"][0]).to.deep.equal({
                "w:tcBorders": [
                    {
                        "w:top": {
                            _attr: { "w:color": "ff0000", "w:sz": 3, "w:val": "dashDotStroked" },
                        },
                    },
                ],
            });
        });

        it("should set the margins", () => {
            const properties = new TableCellProperties({
                margins: {
                    bottom: 15,
                    left: 10,
                    marginUnitType: WidthType.DXA,
                    right: 20,
                    top: 5,
                },
            });

            const tree = new Formatter().format(properties);

            expect(tree["w:tcPr"][0]).to.deep.equal({
                "w:tcMar": [
                    { "w:top": { _attr: { "w:type": "dxa", "w:w": 5 } } },
                    { "w:left": { _attr: { "w:type": "dxa", "w:w": 10 } } },
                    { "w:bottom": { _attr: { "w:type": "dxa", "w:w": 15 } } },
                    { "w:right": { _attr: { "w:type": "dxa", "w:w": 20 } } },
                ],
            });
        });

        it("should not add margins when all margin values are undefined", () => {
            const properties = new TableCellProperties({
                margins: {},
            });
            expect(() => new Formatter().format(properties)).to.throw(
                "XMLComponent did not format correctly",
            );
        });

        it("sets cellIns to track cell insertion", () => {
            const cellProperties = new TableCellProperties({
                insertion: { author: "Firstname Lastname", date: "123", id: 1 },
            });
            const tree = new Formatter().format(cellProperties);
            expect(tree).to.deep.equal({
                "w:tcPr": [
                    {
                        "w:cellIns": {
                            _attr: { "w:author": "Firstname Lastname", "w:date": "123", "w:id": 1 },
                        },
                    },
                ],
            });
        });

        it("sets cellDel to track cell deletion", () => {
            const cellProperties = new TableCellProperties({
                deletion: { author: "Firstname Lastname", date: "123", id: 1 },
            });
            const tree = new Formatter().format(cellProperties);
            expect(tree).to.deep.equal({
                "w:tcPr": [
                    {
                        "w:cellDel": {
                            _attr: { "w:author": "Firstname Lastname", "w:date": "123", "w:id": 1 },
                        },
                    },
                ],
            });
        });

        it("sets cellMerge to track vertical merge revision", () => {
            const cellProperties = new TableCellProperties({
                cellMerge: {
                    author: "Firstname Lastname",
                    date: "123",
                    id: 1,
                    verticalMerge: "rest",
                },
            });
            const tree = new Formatter().format(cellProperties);
            expect(tree).to.deep.equal({
                "w:tcPr": [
                    {
                        "w:cellMerge": {
                            _attr: {
                                "w:author": "Firstname Lastname",
                                "w:date": "123",
                                "w:id": 1,
                                "w:vMerge": "rest",
                            },
                        },
                    },
                ],
            });
        });

        it("should add a revision property", () => {
            const cellProperties = new TableCellProperties({
                revision: {
                    author: "Firstname Lastname",
                    date: "123",
                    id: 1,
                },
            });
            const tree = new Formatter().format(cellProperties);
            expect(tree).to.deep.equal({
                "w:tcPr": [
                    {
                        "w:tcPrChange": [
                            {
                                _attr: {
                                    "w:author": "Firstname Lastname",
                                    "w:date": "123",
                                    "w:id": 1,
                                },
                            },
                            {
                                "w:tcPr": {},
                            },
                        ],
                    },
                ],
            });
        });
    });
});
