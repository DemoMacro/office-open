import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createMathLimitLocation } from "./math-limit-location";

describe("createMathLimitLocation", () => {
    describe("#constructor()", () => {
        it("should create a MathLimitLocation with correct root key", () => {
            const mathLimitLocation = createMathLimitLocation({});

            const tree = new Formatter().format(mathLimitLocation);
            expect(tree).to.deep.equal({
                "m:limLoc": {
                    _attr: {
                        "m:val": "undOvr",
                    },
                },
            });
        });
    });
});
