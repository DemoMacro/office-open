import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { LatentStyleException } from "./exceptions";

describe("LatentStyleException", () => {
    describe("#constructor()", () => {
        it("should create", () => {
            const currentLatentStyleException = new LatentStyleException({
                name: "test-name",
                qFormat: "test-qFormat",
                semiHidden: "test-semiHidden",
                uiPriority: "test-uiPriority",
                unhideWhenUsed: "test-unhideWhenUsed",
            });

            const tree = new Formatter().format(currentLatentStyleException);
            expect(tree).to.deep.equal({
                "w:lsdException": {
                    _attr: {
                        "w:name": "test-name",
                        "w:qFormat": "test-qFormat",
                        "w:semiHidden": "test-semiHidden",
                        "w:uiPriority": "test-uiPriority",
                        "w:unhideWhenUsed": "test-unhideWhenUsed",
                    },
                },
            });
        });
    });
});
