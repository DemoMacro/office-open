import type { IContext } from "@file/xml-components";
import * as convenienceFunctions from "@util/convenience-functions";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { AltChunk } from "./alt-chunk";

describe("AltChunk", () => {
    beforeEach(() => {
        vi.spyOn(convenienceFunctions, "uniqueId").mockReturnValue("1");
    });

    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("should wrap HTML fragment in a full document", () => {
        const altChunk = new AltChunk({
            data: "<p>HTML content</p>",
            contentType: "text/html",
            extension: "html",
        });

        const addRelationship = vi.fn();
        const addAltChunk = vi.fn();
        const addAltChunkContentType = vi.fn();
        const mockContext = {
            file: {
                AltChunks: { addAltChunk },
                ContentTypes: { addAltChunk: addAltChunkContentType },
            },
            viewWrapper: {
                Relationships: { addRelationship },
            },
            stack: [],
        } as unknown as IContext;

        const result = altChunk.prepForXml(mockContext);

        expect(result).to.deep.equal({
            "w:altChunk": {
                _attr: {
                    "r:id": "rId1",
                },
            },
        });

        expect(addRelationship).toHaveBeenCalledWith(
            "1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
            "afchunks/afchunk1.html",
        );

        // HTML fragment should be wrapped in a full document
        const wrappedHtml =
            '<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body><p>HTML content</p></body></html>';
        expect(addAltChunk).toHaveBeenCalledWith("1", {
            key: "1",
            data: new TextEncoder().encode(wrappedHtml),
            path: "afchunks/afchunk1.html",
            extension: "html",
            contentType: "text/html",
        });

        expect(addAltChunkContentType).toHaveBeenCalledWith(
            "/word/afchunks/afchunk1.html",
            "text/html",
            "html",
        );
    });

    it("should not double-wrap an already complete HTML document", () => {
        const altChunk = new AltChunk({
            data: '<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body><p>Already full</p></body></html>',
            contentType: "text/html",
            extension: "html",
        });

        const addAltChunk = vi.fn();
        const mockContext = {
            file: {
                AltChunks: { addAltChunk },
                ContentTypes: { addAltChunk: vi.fn() },
            },
            viewWrapper: {
                Relationships: { addRelationship: vi.fn() },
            },
            stack: [],
        } as unknown as IContext;

        altChunk.prepForXml(mockContext);

        // Data should remain unchanged (no double wrapping)
        const originalHtml =
            '<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body><p>Already full</p></body></html>';
        expect(addAltChunk).toHaveBeenCalledWith(
            "1",
            expect.objectContaining({
                data: new TextEncoder().encode(originalHtml),
            }),
        );
    });

    it("should not wrap RTF content", () => {
        const altChunk = new AltChunk({
            data: "{\\rtf1 Hello}",
            contentType: "application/rtf",
            extension: "rtf",
        });

        const addAltChunk = vi.fn();
        const mockContext = {
            file: {
                AltChunks: { addAltChunk },
                ContentTypes: { addAltChunk: vi.fn() },
            },
            viewWrapper: {
                Relationships: { addRelationship: vi.fn() },
            },
            stack: [],
        } as unknown as IContext;

        altChunk.prepForXml(mockContext);

        // RTF should remain unchanged
        expect(addAltChunk).toHaveBeenCalledWith(
            "1",
            expect.objectContaining({
                data: new TextEncoder().encode("{\\rtf1 Hello}"),
            }),
        );
    });

    it("should create an altChunk element with matchSrc", () => {
        const altChunk = new AltChunk({
            data: "<p>HTML content</p>",
            contentType: "text/html",
            extension: "html",
            matchSrc: true,
        });

        const mockContext = {
            file: {
                AltChunks: { addAltChunk: vi.fn() },
                ContentTypes: { addAltChunk: vi.fn() },
            },
            viewWrapper: {
                Relationships: { addRelationship: vi.fn() },
            },
            stack: [],
        } as unknown as IContext;

        const result = altChunk.prepForXml(mockContext);

        expect(result).to.deep.equal({
            "w:altChunk": [
                {
                    _attr: {
                        "r:id": "rId1",
                    },
                },
                {
                    "w:altChunkPr": [
                        {
                            "w:matchSrc": {},
                        },
                    ],
                },
            ],
        });
    });
});
