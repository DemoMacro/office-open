import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import {
    PositionalTab,
    PositionalTabAlignment,
    PositionalTabLeader,
    PositionalTabRelativeTo,
} from "./positional-tab";

describe("PositionalTab", () => {
    it("should create a PositionalTab with correct root key", () => {
        const tree = new Formatter().format(
            new PositionalTab({
                alignment: PositionalTabAlignment.CENTER,
                leader: PositionalTabLeader.DOT,
                relativeTo: PositionalTabRelativeTo.MARGIN,
            }),
        );

        expect(tree).to.deep.equal({
            "w:ptab": {
                _attr: {
                    "w:alignment": "center",
                    "w:leader": "dot",
                    "w:relativeTo": "margin",
                },
            },
        });
    });
});
