import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { InsertedTextRun } from "./inserted-text-run";

describe("InsertedTextRun", () => {
    describe("#constructor", () => {
        it("should create a inserted text run", () => {
            const insertedTextRun = new InsertedTextRun({
                author: "Author",
                date: "123",
                id: 0,
                text: "some text",
            });
            const tree = new Formatter().format(insertedTextRun);
            expect(tree).to.deep.equal({
                "w:ins": [
                    {
                        _attr: {
                            "w:author": "Author",
                            "w:date": "123",
                            "w:id": 0,
                        },
                    },
                    {
                        "w:r": [
                            {
                                "w:t": [
                                    {
                                        _attr: {
                                            "xml:space": "preserve",
                                        },
                                    },
                                    "some text",
                                ],
                            },
                        ],
                    },
                ],
            });
        });
    });
});
