import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { InsertedTableCell } from "./inserted-table-cell";

describe("InsertedTableCell", () => {
    describe("#constructor", () => {
        it("should create the insertion for table cell", () => {
            const insertion = new InsertedTableCell({ author: "Author", date: "123", id: 0 });
            const tree = new Formatter().format(insertion);
            expect(tree).to.deep.equal({
                "w:cellIns": {
                    _attr: {
                        "w:author": "Author",
                        "w:date": "123",
                        "w:id": 0,
                    },
                },
            });
        });
    });
});
