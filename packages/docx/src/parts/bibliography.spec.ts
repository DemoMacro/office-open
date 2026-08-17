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
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return bibliographyDesc.parse(el, readCtx);
}

describe("bibliographyDesc round-trip", () => {
  it("round-trips single source with basic fields", () => {
    const result = roundTrip({
      sources: [
        {
          sourceType: "Book",
          title: "TypeScript in Action",
          author: { authors: [{ last: "Doe", first: "John" }] },
        },
      ],
    });
    expect(result.sources).toHaveLength(1);
    expect(result.sources[0]?.sourceType).toBe("Book");
    expect(result.sources[0]?.title).toBe("TypeScript in Action");
    expect(result.sources[0]?.author).toEqual({ authors: [{ last: "Doe", first: "John" }] });
  });

  it("round-trips styleName", () => {
    const result = roundTrip({
      sources: [],
      styleName: "APA",
    });
    expect(result.styleName).toBe("APA");
  });

  it("round-trips structured authors across roles", () => {
    const author = {
      authors: [{ corporate: "Microsoft" }, { last: "Smith", first: "J.", middle: "R." }],
      editors: [{ last: "Lee", first: "Kai" }],
      translators: [{ last: "Garcia" }],
    };
    const result = roundTrip({ sources: [{ title: "Film Study", author }] });
    expect(result.sources[0]?.author).toEqual(author);
  });

  it("emits CT_AuthorType role structure, not flat text", () => {
    const xml = bibliographyDesc.stringify(
      {
        sources: [
          {
            title: "Doc",
            author: {
              authors: [{ last: "Doe", first: "Jane" }],
              editors: [{ last: "Lee", first: "Kai" }],
            },
          },
        ],
      },
      writeCtx,
    )!;
    expect(xml).toContain(
      "<b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Jane</b:First>" +
        "</b:Person></b:NameList></b:Author>",
    );
    expect(xml).toContain(
      "<b:Editor><b:NameList><b:Person><b:Last>Lee</b:Last><b:First>Kai</b:First></b:Person></b:NameList></b:Editor>",
    );
  });

  it("round-trips a corporate author entry", () => {
    const result = roundTrip({
      sources: [{ title: "Docs", author: { authors: [{ corporate: "Microsoft" }] } }],
    });
    expect(result.sources[0]?.author?.authors).toEqual([{ corporate: "Microsoft" }]);
  });

  it("round-trips all source fields", () => {
    const result = roundTrip({
      sources: [
        {
          sourceType: "JournalArticle",
          title: "Deep Learning",
          author: { authors: [{ last: "Smith", first: "Jane" }] },
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
    if (!src) throw new Error("source not parsed");
    expect(src.sourceType).toBe("JournalArticle");
    expect(src.title).toBe("Deep Learning");
    expect(src.author).toEqual({ authors: [{ last: "Smith", first: "Jane" }] });
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
    expect(result.sources[0]?.title).toBe("First");
    expect(result.sources[1]?.title).toBe("Second");
    expect(result.sources[2]?.title).toBe("Third");
  });

  it("round-trips empty sources", () => {
    const result = roundTrip({ sources: [] });
    expect(result.sources).toHaveLength(0);
  });

  it("handles XML special characters", () => {
    const result = roundTrip({
      sources: [{ title: 'Tom & Jerry "The Movie"' }],
    });
    expect(result.sources[0]?.title).toBe('Tom & Jerry "The Movie"');
  });
});
