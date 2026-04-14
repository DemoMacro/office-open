import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { LeaderType, TabStopType, createTabStop } from "./tab-stop";

describe("LeftTabStop", () => {
    describe("#createTabStop()", () => {
        it("should create a Tab Stop with correct attributes", () => {
            const tabStop = createTabStop([{ position: 100, type: TabStopType.LEFT }]);
            const tree = new Formatter().format(tabStop);
            expect(tree).to.deep.equal({
                "w:tabs": [
                    {
                        "w:tab": {
                            _attr: {
                                "w:pos": 100,
                                "w:val": "left",
                            },
                        },
                    },
                ],
            });
        });
    });
});

describe("RightTabStop", () => {
    describe("#createTabStop()", () => {
        it("should create a Tab Stop with correct attributes", () => {
            const tabStop = createTabStop([
                { leader: LeaderType.DOT, position: 100, type: TabStopType.RIGHT },
            ]);
            const tree = new Formatter().format(tabStop);
            expect(tree).to.deep.equal({
                "w:tabs": [
                    {
                        "w:tab": {
                            _attr: {
                                "w:leader": "dot",
                                "w:pos": 100,
                                "w:val": "right",
                            },
                        },
                    },
                ],
            });
        });
    });
});
