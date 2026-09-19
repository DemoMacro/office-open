import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument, generateDocumentStream, generateDocumentSync } from "./generate";
import type { DocumentOptions } from "./parts/core-properties";

const DATE = "2024-05-06T07:08:09.000Z";
const OTHER_DATE = "2025-02-03T04:05:06.000Z";

/** Fonts and comments exercise the scope-driven id and date readers. */
const options = (): DocumentOptions => ({
  sections: [{ children: [{ paragraph: { children: ["Reproducible"] } }] }],
  fonts: [{ name: "ReproFont", data: new Uint8Array(64).fill(3) }],
  comments: [{ id: 1, author: "A", children: [{ children: ["note"] }] }],
});

const decode = (buffer: Uint8Array, path: string): string =>
  new TextDecoder().decode(unzipSync(buffer)[path]);

const collect = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> => {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (value) chunks.push(value);
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
};

const generate = {
  async: (date: string) => generateDocument(options(), { reproducible: { date } }),
  sync: (date: string) => generateDocumentSync(options(), { reproducible: { date } }),
  stream: (date: string) => collect(generateDocumentStream(options(), { reproducible: { date } })),
};

/** Revision markers of every fresh-constructible kind exercise the revision-id
 *  threading (autoRevisionId) — a leaked module-counter fallback shows up as a
 *  byte difference because the plain generation below advances the counters. */
const revisionOptions = (): DocumentOptions => ({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              { insertion: { author: "A", date: DATE, children: [{ text: "ins" }] } },
              { deletion: { author: "A", date: DATE, children: [{ text: "del" }] } },
              {
                moveFrom: { name: "m1", author: "A", date: DATE, wrap: ["mv"] },
              },
              {
                moveTo: { name: "m1", author: "A", date: DATE, wrap: ["mv"] },
              },
            ],
          },
        },
        {
          paragraph: {
            children: ["para"],
            revision: { author: "A", date: DATE, alignment: "center" },
          },
        },
        {
          table: {
            revision: { author: "A", date: DATE },
            rows: [{ cells: [{ children: [{ paragraph: { children: ["cell"] } }] }] }],
          },
        },
      ],
    },
  ],
});

describe("reproducible generation", () => {
  it("derives revision marker ids from the reproducible option", () => {
    // Advance the module-level fallback counters first — a leaked fallback
    // read would then produce a byte difference between the two scoped runs.
    generateDocumentSync(revisionOptions());
    const first = generateDocumentSync(revisionOptions(), { reproducible: { date: DATE } });
    const second = generateDocumentSync(revisionOptions(), { reproducible: { date: DATE } });
    expect(Array.from(first)).toEqual(Array.from(second));
    const doc = decode(first, "word/document.xml");
    expect(doc).toContain("<w:ins ");
    expect(doc).toContain("<w:del ");
    expect(doc).toContain("<w:moveFrom ");
    expect(doc).toContain("<w:pPrChange ");
    expect(doc).toContain("<w:tblPrChange ");
  });

  it("generates byte-identical documents on the async, sync and stream paths", async () => {
    const async_ = await generate.async(DATE);
    const sync = generate.sync(DATE);
    const stream = await generate.stream(DATE);
    // Byte equality across the Buffer (async/sync) and Uint8Array (stream) types.
    expect(Array.from(async_)).toEqual(Array.from(sync));
    expect(Array.from(stream)).toEqual(Array.from(sync));
  });

  it("drives the date defaults from the reproducible option", () => {
    const buffer = generate.sync(DATE);
    expect(decode(buffer, "word/comments.xml")).toContain(`w:date="${DATE}"`);
    expect(decode(buffer, "docProps/core.xml")).toContain(
      `<dcterms:created xsi:type="dcterms:W3CDTF">${DATE}`,
    );
    expect(decode(buffer, "docProps/core.xml")).toContain(
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${DATE}`,
    );
  });

  it("isolates concurrent generations from each other", async () => {
    const soloFirst = await generate.async(DATE);
    const soloSecond = await generate.async(OTHER_DATE);

    // Each call creates its own scope inside the packer — no shared state.
    const [first, second] = await Promise.all([generate.async(DATE), generate.async(OTHER_DATE)]);
    expect(first).toEqual(soloFirst);
    expect(second).toEqual(soloSecond);
    expect(first).not.toEqual(second);
  });

  it("keeps cryptographic ids and real dates without the reproducible option", () => {
    const first = generateDocumentSync(options());
    const second = generateDocumentSync(options());
    expect(first).not.toEqual(second);
    expect(decode(first, "docProps/core.xml")).not.toContain(DATE);
  });
});
