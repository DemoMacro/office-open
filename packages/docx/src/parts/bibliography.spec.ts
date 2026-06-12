import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { bibliographyDesc } from "./bibliography";
import type { BibliographyOptions } from "./bibliography";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: BibliographyOptions) {
  const xml = bibliographyDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return bibliographyDesc.parse(el, readCtx);
}

describe("bibliographyDesc round-trip", () => {
  it("round-trips single source with basic fields", () => {
    const result = roundTrip({
      sources: [{ type: "Book", title: "TypeScript in Action", author: "John Doe" }],
    });
    expect(result.sources).toHaveLength(1);
    expect(result.sources[0].type).toBe("Book");
    expect(result.sources[0].title).toBe("TypeScript in Action");
    expect(result.sources[0].author).toBe("John Doe");
  });

  it("round-trips styleName", () => {
    const result = roundTrip({
      sources: [],
      styleName: "APA",
    });
    expect(result.styleName).toBe("APA");
  });

  it("round-trips all source fields", () => {
    const result = roundTrip({
      sources: [
        {
          type: "JournalArticle",
          title: "Deep Learning",
          author: "Jane Smith",
          year: "2024",
          month: "03",
          day: "15",
          bookTitle: "AI Handbook",
          journal: "Nature AI",
          volume: "12",
          issue: "4",
          pages: "100-120",
          publisher: "Springer",
          city: "Berlin",
          url: "https://example.com",
          edition: "2nd",
          institution: "MIT",
        },
      ],
    });
    const src = result.sources[0];
    expect(src.type).toBe("JournalArticle");
    expect(src.title).toBe("Deep Learning");
    expect(src.author).toBe("Jane Smith");
    expect(src.year).toBe("2024");
    expect(src.month).toBe("03");
    expect(src.day).toBe("15");
    expect(src.bookTitle).toBe("AI Handbook");
    expect(src.journal).toBe("Nature AI");
    expect(src.volume).toBe("12");
    expect(src.issue).toBe("4");
    expect(src.pages).toBe("100-120");
    expect(src.publisher).toBe("Springer");
    expect(src.city).toBe("Berlin");
    expect(src.url).toBe("https://example.com");
    expect(src.edition).toBe("2nd");
    expect(src.institution).toBe("MIT");
  });

  it("round-trips multiple sources", () => {
    const result = roundTrip({
      sources: [{ title: "First" }, { title: "Second" }, { title: "Third" }],
    });
    expect(result.sources).toHaveLength(3);
    expect(result.sources[0].title).toBe("First");
    expect(result.sources[1].title).toBe("Second");
    expect(result.sources[2].title).toBe("Third");
  });

  it("round-trips empty sources", () => {
    const result = roundTrip({ sources: [] });
    expect(result.sources).toHaveLength(0);
  });

  it("handles XML special characters", () => {
    const result = roundTrip({
      sources: [{ title: 'Tom & Jerry "The Movie"' }],
    });
    expect(result.sources[0].title).toBe('Tom & Jerry "The Movie"');
  });
});
