import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createLanguageComponent } from "./language";

describe("Language", () => {
    describe("#createLanguageComponent", () => {
        it("should create a language component", () => {
            const tree = new Formatter().format(
                createLanguageComponent({
                    bidirectional: "ar-SA",
                    eastAsia: "zh-CN",
                    value: "en-US",
                }),
            );

            expect(tree).to.deep.equal({
                "w:lang": {
                    _attr: {
                        "w:bidi": "ar-SA",
                        "w:eastAsia": "zh-CN",
                        "w:val": "en-US",
                    },
                },
            });
        });
    });
});
