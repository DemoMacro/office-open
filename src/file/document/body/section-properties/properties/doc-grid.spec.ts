import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { DocumentGridType, createDocumentGrid } from ".";

describe("createDocumentGrid", () => {
    describe("#constructor()", () => {
        it("should create documentGrid with specified linePitch", () => {
            const docGrid = createDocumentGrid({ linePitch: 360 });
            const tree = new Formatter().format(docGrid);

            expect(tree["w:docGrid"]).to.deep.equal({ _attr: { "w:linePitch": 360 } });
        });

        it("should create documentGrid with specified linePitch and type", () => {
            const docGrid = createDocumentGrid({ linePitch: 360, type: DocumentGridType.LINES });
            const tree = new Formatter().format(docGrid);

            expect(tree["w:docGrid"]).to.deep.equal({
                _attr: { "w:linePitch": 360, "w:type": "lines" },
            });
        });

        it("should create documentGrid with specified linePitch,charSpace and type", () => {
            const docGrid = createDocumentGrid({
                charSpace: -1541,
                linePitch: 346,
                type: DocumentGridType.LINES_AND_CHARS,
            });
            const tree = new Formatter().format(docGrid);

            expect(tree["w:docGrid"]).to.deep.equal({
                _attr: { "w:charSpace": -1541, "w:linePitch": 346, "w:type": "linesAndChars" },
            });
        });
    });
});
