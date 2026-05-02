import type { IContext } from "@file/xml-components";
import * as convenienceFunctions from "@util/convenience-functions";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { SubDoc } from "./sub-doc";

describe("SubDoc", () => {
    beforeEach(() => {
        vi.spyOn(convenienceFunctions, "uniqueId").mockReturnValue("1");
    });

    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("should create a subDoc element with r:id", () => {
        const subDoc = new SubDoc({
            data: new Uint8Array([0x50, 0x4b]),
        });

        const addRelationship = vi.fn();
        const addSubDoc = vi.fn();
        const addSubDocContentType = vi.fn();
        const mockContext = {
            file: {
                SubDocs: { addSubDoc },
                ContentTypes: { addSubDoc: addSubDocContentType },
            },
            viewWrapper: {
                Relationships: { addRelationship },
            },
            stack: [],
        } as unknown as IContext;

        const result = subDoc.prepForXml(mockContext);

        expect(result).to.deep.equal({
            "w:subDoc": {
                _attr: {
                    "r:id": "rId1",
                },
            },
        });

        expect(addRelationship).toHaveBeenCalledWith(
            "1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument",
            "subdocs/subdoc1.docx",
        );

        expect(addSubDoc).toHaveBeenCalledWith("1", {
            data: new Uint8Array([0x50, 0x4b]),
            path: "subdocs/subdoc1.docx",
        });

        expect(addSubDocContentType).toHaveBeenCalledWith("/word/subdocs/subdoc1.docx");
    });

    it("should accept string data and convert to bytes", () => {
        const subDoc = new SubDoc({
            data: "test-docx-content",
        });

        const addSubDoc = vi.fn();
        const mockContext = {
            file: {
                SubDocs: { addSubDoc },
                ContentTypes: { addSubDoc: vi.fn() },
            },
            viewWrapper: {
                Relationships: { addRelationship: vi.fn() },
            },
            stack: [],
        } as unknown as IContext;

        subDoc.prepForXml(mockContext);

        expect(addSubDoc).toHaveBeenCalledWith(
            "1",
            expect.objectContaining({
                data: new TextEncoder().encode("test-docx-content"),
            }),
        );
    });
});
