import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Bibliography } from "./bibliography";

describe("Bibliography", () => {
    it("should create a bibliography with no sources", () => {
        const tree = new Formatter().format(new Bibliography({ sources: [] }));
        expect(tree).to.deep.equal({
            "b:Sources": {
                _attr: {
                    "xmlns:b": "http://purl.oclc.org/ooxml/officeDocument/bibliography",
                },
            },
        });
    });

    it("should create a bibliography with a book source", () => {
        const tree = new Formatter().format(
            new Bibliography({
                styleName: "APA",
                sources: [
                    {
                        type: "Book",
                        title: "The Design of Everyday Things",
                        author: "Norman, Donald",
                        year: "2013",
                        publisher: "Basic Books",
                        city: "New York",
                        edition: "Revised",
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "b:Sources": [
                {
                    _attr: {
                        "xmlns:b": "http://purl.oclc.org/ooxml/officeDocument/bibliography",
                        StyleName: "APA",
                    },
                },
                {
                    "b:Source": [
                        { "b:SourceType": ["Book"] },
                        { "b:Title": ["The Design of Everyday Things"] },
                        { "b:Author": ["Norman, Donald"] },
                        { "b:Year": ["2013"] },
                        { "b:Publisher": ["Basic Books"] },
                        { "b:City": ["New York"] },
                        { "b:Edition": ["Revised"] },
                    ],
                },
            ],
        });
    });

    it("should create a bibliography with a journal article source", () => {
        const tree = new Formatter().format(
            new Bibliography({
                sources: [
                    {
                        type: "JournalArticle",
                        title: "A Survey of Techniques",
                        author: "Smith, J.; Doe, A.",
                        year: "2026",
                        month: "April",
                        journal: "Journal of Examples",
                        volume: "42",
                        issue: "3",
                        pages: "100-120",
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "b:Sources": [
                {
                    _attr: {
                        "xmlns:b": "http://purl.oclc.org/ooxml/officeDocument/bibliography",
                    },
                },
                {
                    "b:Source": [
                        { "b:SourceType": ["JournalArticle"] },
                        { "b:Title": ["A Survey of Techniques"] },
                        { "b:Author": ["Smith, J.; Doe, A."] },
                        { "b:Year": ["2026"] },
                        { "b:Month": ["April"] },
                        { "b:JournalName": ["Journal of Examples"] },
                        { "b:Volume": ["42"] },
                        { "b:Issue": ["3"] },
                        { "b:Pages": ["100-120"] },
                    ],
                },
            ],
        });
    });

    it("should create a bibliography with multiple sources", () => {
        const tree = new Formatter().format(
            new Bibliography({
                sources: [
                    {
                        type: "Book",
                        title: "First Book",
                        year: "2020",
                    },
                    {
                        type: "InternetSite",
                        title: "Online Resource",
                        url: "https://example.com",
                    },
                ],
            }),
        );
        // Verify both sources are present
        const sources = tree["b:Sources"] as unknown[];
        expect(sources).to.have.length(3); // attributes + 2 sources
        expect(sources[1]).to.deep.equal({
            "b:Source": [
                { "b:SourceType": ["Book"] },
                { "b:Title": ["First Book"] },
                { "b:Year": ["2020"] },
            ],
        });
        expect(sources[2]).to.deep.equal({
            "b:Source": [
                { "b:SourceType": ["InternetSite"] },
                { "b:Title": ["Online Resource"] },
                { "b:URL": ["https://example.com"] },
            ],
        });
    });

    it("should expose a Relationships instance", () => {
        const bib = new Bibliography({ sources: [] });
        expect(bib.Relationships).to.be.ok;
        expect(bib.Relationships.RelationshipCount).to.equal(0);
    });
});
