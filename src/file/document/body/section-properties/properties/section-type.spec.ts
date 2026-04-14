import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { SectionType, createSectionType } from "./section-type";

describe("createSectionType", () => {
    it("should create with even page section type", () => {
        const sectionType = createSectionType(SectionType.EVEN_PAGE);

        const tree = new Formatter().format(sectionType);

        expect(tree).to.deep.equal({
            "w:type": {
                _attr: {
                    "w:val": "evenPage",
                },
            },
        });
    });

    it("should create with continuous section type", () => {
        const sectionType = createSectionType(SectionType.CONTINUOUS);

        const tree = new Formatter().format(sectionType);

        expect(tree).to.deep.equal({
            "w:type": {
                _attr: {
                    "w:val": "continuous",
                },
            },
        });
    });
});
