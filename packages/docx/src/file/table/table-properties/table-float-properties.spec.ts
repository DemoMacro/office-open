import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import {
    OverlapType,
    RelativeHorizontalPosition,
    RelativeVerticalPosition,
    TableAnchorType,
    createTableFloatProperties,
} from "./table-float-properties";

describe("Table Float Properties", () => {
    describe("#createTableFloatProperties", () => {
        it("should construct a TableFloatProperties with all options", () => {
            const properties = createTableFloatProperties({
                absoluteHorizontalPosition: 10,
                absoluteVerticalPosition: 20,
                bottomFromText: 30,
                horizontalAnchor: TableAnchorType.MARGIN,
                leftFromText: 50,
                relativeHorizontalPosition: RelativeHorizontalPosition.CENTER,
                relativeVerticalPosition: RelativeVerticalPosition.BOTTOM,
                rightFromText: 60,
                topFromText: 40,
                verticalAnchor: TableAnchorType.PAGE,
            });
            const tree = new Formatter().format(properties);
            expect(tree).to.deep.equal(DEFAULT_TFP);
        });

        it("should add overlap", () => {
            const properties = createTableFloatProperties({
                overlap: OverlapType.NEVER,
            });
            const tree = new Formatter().format(properties);

            expect(tree).to.deep.equal({
                "w:tblpPr": {
                    _attr: {},
                },
            });
        });
    });
});

const DEFAULT_TFP = {
    "w:tblpPr": {
        _attr: {
            "w:bottomFromText": 30,
            "w:horzAnchor": "margin",
            "w:leftFromText": 50,
            "w:rightFromText": 60,
            "w:tblpX": 10,
            "w:tblpXSpec": "center",
            "w:tblpY": 20,
            "w:tblpYSpec": "bottom",
            "w:topFromText": 40,
            "w:vertAnchor": "page",
        },
    },
};
