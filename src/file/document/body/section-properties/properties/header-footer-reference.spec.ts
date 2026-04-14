import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import {
    HeaderFooterReferenceType,
    HeaderFooterType,
    createHeaderFooterReference,
} from "./header-footer-reference";

describe("createHeaderFooterReference", () => {
    it("should create footer reference", () => {
        const footer = createHeaderFooterReference(HeaderFooterType.FOOTER, {
            id: 1,
            type: HeaderFooterReferenceType.DEFAULT,
        });

        const tree = new Formatter().format(footer);
        expect(tree).to.deep.equal({
            "w:footerReference": {
                _attr: {
                    "r:id": "rId1",
                    "w:type": "default",
                },
            },
        });
    });

    it("should create header reference", () => {
        const header = createHeaderFooterReference(HeaderFooterType.HEADER, {
            id: 1,
            type: HeaderFooterReferenceType.DEFAULT,
        });

        const tree = new Formatter().format(header);
        expect(tree).to.deep.equal({
            "w:headerReference": {
                _attr: {
                    "r:id": "rId1",
                    "w:type": "default",
                },
            },
        });
    });

    it("should create without a type (defaults to DEFAULT)", () => {
        const footer = createHeaderFooterReference(HeaderFooterType.FOOTER, {
            id: 1,
        });

        const tree = new Formatter().format(footer);
        expect(tree).to.deep.equal({
            "w:footerReference": {
                _attr: {
                    "r:id": "rId1",
                    "w:type": "default",
                },
            },
        });
    });
});
