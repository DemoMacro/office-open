import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { CellMerge, VerticalMergeRevisionType } from "./cell-merge";

describe("CellMerge", () => {
    describe("#constructor()", () => {
        it("creates with vMerge attribute", () => {
            const tree = new Formatter().format(
                new CellMerge({
                    author: "Firstname Lastname",
                    date: "123",
                    id: 1,
                    verticalMerge: VerticalMergeRevisionType.CONTINUE,
                }),
            );
            expect(tree).to.deep.equal({
                "w:cellMerge": {
                    _attr: {
                        "w:author": "Firstname Lastname",
                        "w:date": "123",
                        "w:id": 1,
                        "w:vMerge": "cont",
                    },
                },
            });
        });

        it("creates with vMergeOrig attribute", () => {
            const tree = new Formatter().format(
                new CellMerge({
                    author: "Firstname Lastname",
                    date: "123",
                    id: 1,
                    verticalMergeOriginal: VerticalMergeRevisionType.RESTART,
                }),
            );
            expect(tree).to.deep.equal({
                "w:cellMerge": {
                    _attr: {
                        "w:author": "Firstname Lastname",
                        "w:date": "123",
                        "w:id": 1,
                        "w:vMergeOrig": "rest",
                    },
                },
            });
        });
    });
});
