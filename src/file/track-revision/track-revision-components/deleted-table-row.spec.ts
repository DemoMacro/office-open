import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { DeletedTableRow } from "./deleted-table-row";

describe("DeletedTableRow", () => {
    describe("#constructor", () => {
        it("should create the deletion for table row", () => {
            const deletion = new DeletedTableRow({ author: "Author", date: "123", id: 0 });
            const tree = new Formatter().format(deletion);
            expect(tree).to.deep.equal({
                "w:del": {
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
