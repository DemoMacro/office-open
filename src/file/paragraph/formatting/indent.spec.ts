import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createIndent } from "./indent";

describe("Indent", () => {
    it("should create", () => {
        const indent = createIndent({
            end: 10,
            firstLine: 10,
            hanging: 10,
            left: 10,
            right: 10,
            start: 10,
        });
        const tree = new Formatter().format(indent);
        expect(tree).to.deep.equal({
            "w:ind": {
                _attr: {
                    "w:end": 10,
                    "w:firstLine": 10,
                    "w:hanging": 10,
                    "w:left": 10,
                    "w:right": 10,
                    "w:start": 10,
                },
            },
        });
    });

    it("should create with no indent values", () => {
        const indent = createIndent({});

        const tree = new Formatter().format(indent);
        expect(tree).to.deep.equal({
            "w:ind": {
                _attr: {},
            },
        });
    });
});
