import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { InsertedTableRow } from "./inserted-table-row";

describe("InsertedTableRow", () => {
    describe("#constructor", () => {
        it("should create the insertion for table row", () => {
            const insertion = new InsertedTableRow({ author: "Author", date: "123", id: 0 });
            const tree = new Formatter().format(insertion);
            expect(tree).to.deep.equal({
                "w:ins": {
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
