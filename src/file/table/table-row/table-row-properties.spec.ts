import { Formatter } from "@export/formatter";
import { AlignmentType } from "@file/paragraph";
import { HeightRule } from "@file/table/table-row/table-row-height";
import { describe, expect, it } from "vite-plus/test";

import { CellSpacingType } from "../table-cell-spacing";
import { WidthType } from "../table-width";
import { TableRowProperties } from "./table-row-properties";

describe("TableRowProperties", () => {
    describe("#constructor", () => {
        it("creates an initially empty property object", () => {
            const rowProperties = new TableRowProperties({});
            // The TableRowProperties is ignorable if there are no attributes,
            // Which results in prepForXml returning undefined, which causes
            // The formatter to throw an error if that is the only object it
            // Has been asked to format.
            expect(() => new Formatter().format(rowProperties)).to.throw(
                "XMLComponent did not format correctly",
            );
        });

        it("sets cantSplit to avoid row been paginated", () => {
            const rowProperties = new TableRowProperties({ cantSplit: true });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({ "w:trPr": [{ "w:cantSplit": {} }] });
        });

        it("sets ins to track row insertion", () => {
            const rowProperties = new TableRowProperties({
                insertion: { author: "Firstname Lastname", date: "123", id: 1 },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [
                    {
                        "w:ins": {
                            _attr: { "w:author": "Firstname Lastname", "w:date": "123", "w:id": 1 },
                        },
                    },
                ],
            });
        });

        it("sets del to track row deletion", () => {
            const rowProperties = new TableRowProperties({
                deletion: { author: "Firstname Lastname", date: "123", id: 1 },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [
                    {
                        "w:del": {
                            _attr: { "w:author": "Firstname Lastname", "w:date": "123", "w:id": 1 },
                        },
                    },
                ],
            });
        });

        it("sets row as table header (repeat row on each page of table)", () => {
            const rowProperties = new TableRowProperties({ tableHeader: true });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({ "w:trPr": [{ "w:tblHeader": {} }] });
        });

        it("sets row height exact", () => {
            const rowProperties = new TableRowProperties({
                height: {
                    rule: HeightRule.EXACT,
                    value: 100,
                },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:trHeight": { _attr: { "w:hRule": "exact", "w:val": 100 } } }],
            });
        });

        it("sets row height auto", () => {
            const rowProperties = new TableRowProperties({
                height: {
                    rule: HeightRule.AUTO,
                    value: 100,
                },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:trHeight": { _attr: { "w:hRule": "auto", "w:val": 100 } } }],
            });
        });

        it("sets row height at least", () => {
            const rowProperties = new TableRowProperties({
                height: {
                    rule: HeightRule.ATLEAST,
                    value: 100,
                },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:trHeight": { _attr: { "w:hRule": "atLeast", "w:val": 100 } } }],
            });
        });

        it("should add a table cell spacing property", () => {
            const rowProperties = new TableRowProperties({
                cellSpacing: {
                    type: CellSpacingType.DXA,
                    value: 1234,
                },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:tblCellSpacing": { _attr: { "w:type": "dxa", "w:w": 1234 } } }],
            });
        });

        it("should add a revision property", () => {
            const rowProperties = new TableRowProperties({
                revision: {
                    author: "Firstname Lastname",
                    date: "123",
                    id: 1,
                },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [
                    {
                        "w:trPrChange": [
                            {
                                _attr: {
                                    "w:author": "Firstname Lastname",
                                    "w:date": "123",
                                    "w:id": 1,
                                },
                            },
                            {
                                "w:trPr": {},
                            },
                        ],
                    },
                ],
            });
        });

        it("adds cnfStyle with val", () => {
            const rowProperties = new TableRowProperties({
                cnfStyle: { val: "1000000000000010" },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [
                    { "w:cnfStyle": { _attr: { "w:val": "1000000000000010" } } },
                ],
            });
        });

        it("adds cnfStyle with val and changed", () => {
            const rowProperties = new TableRowProperties({
                cnfStyle: { val: "1000000000000010", changed: true },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [
                    { "w:cnfStyle": { _attr: { "w:val": "1000000000000010", "w:changed": true } } },
                ],
            });
        });

        it("adds divId", () => {
            const rowProperties = new TableRowProperties({ divId: 123 });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:divId": { _attr: { "w:val": 123 } } }],
            });
        });

        it("adds gridBefore", () => {
            const rowProperties = new TableRowProperties({ gridBefore: 2 });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:gridBefore": { _attr: { "w:val": 2 } } }],
            });
        });

        it("adds gridAfter", () => {
            const rowProperties = new TableRowProperties({ gridAfter: 1 });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:gridAfter": { _attr: { "w:val": 1 } } }],
            });
        });

        it("adds wBefore", () => {
            const rowProperties = new TableRowProperties({
                widthBefore: { size: 200, type: WidthType.DXA },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:wBefore": { _attr: { "w:type": "dxa", "w:w": 200 } } }],
            });
        });

        it("adds wAfter", () => {
            const rowProperties = new TableRowProperties({
                widthAfter: { size: 300, type: WidthType.DXA },
            });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:wAfter": { _attr: { "w:type": "dxa", "w:w": 300 } } }],
            });
        });

        it("adds row alignment (jc)", () => {
            const rowProperties = new TableRowProperties({ rowAlignment: AlignmentType.CENTER });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:jc": { _attr: { "w:val": "center" } } }],
            });
        });

        it("adds hidden", () => {
            const rowProperties = new TableRowProperties({ hidden: true });
            const tree = new Formatter().format(rowProperties);
            expect(tree).to.deep.equal({
                "w:trPr": [{ "w:hidden": {} }],
            });
        });
    });
});
