import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { DeletedTableCell } from "./deleted-table-cell";

describe("DeletedTableCell", () => {
    describe("#constructor", () => {
        it("should create the deletion for table cell", () => {
            const deletion = new DeletedTableCell({ author: "Author", date: "123", id: 0 });
            const tree = new Formatter().format(deletion);
            expect(tree).to.deep.equal({
                "w:cellDel": {
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
