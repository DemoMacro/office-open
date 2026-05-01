import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { SequentialIdentifier } from "./sequential-identifier";

describe("Sequential Identifier", () => {
    describe("#constructor", () => {
        it("should construct a SEQ without options", () => {
            const seq = new SequentialIdentifier("Figure");
            const tree = new Formatter().format(seq);
            expect(tree).to.be.deep.equal(DEFAULT_SEQ);
        });
    });
});

const DEFAULT_SEQ = {
    "w:r": [
        {
            "w:fldChar": {
                _attr: {
                    "w:dirty": true,
                    "w:fldCharType": "begin",
                },
            },
        },
        {
            "w:instrText": [
                {
                    _attr: {
                        "xml:space": "preserve",
                    },
                },
                "SEQ Figure",
            ],
        },
        {
            "w:fldChar": {
                _attr: {
                    "w:fldCharType": "separate",
                },
            },
        },
        {
            "w:fldChar": {
                _attr: {
                    "w:fldCharType": "end",
                },
            },
        },
    ],
};
