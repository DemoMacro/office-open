import { unzipSync, withReproducibleGeneration } from "@office-open/core";
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

const scoped = <T>(date: string, fn: () => T): T => withReproducibleGeneration({ date }, fn);

describe("reproducible generation", () => {
  it("generates byte-identical documents on the async, sync and stream paths", async () => {
    const async_ = await scoped(DATE, () => generateDocument(options()));
    const sync = scoped(DATE, () => generateDocumentSync(options()));
    const stream = await scoped(DATE, () => collect(generateDocumentStream(options())));
    // Byte equality across the Buffer (async/sync) and Uint8Array (stream) types.
    expect(Array.from(async_)).toEqual(Array.from(sync));
    expect(Array.from(stream)).toEqual(Array.from(sync));
  });

  it("drives the date defaults from the scope", () => {
    const buffer = scoped(DATE, () => generateDocumentSync(options()));
    expect(decode(buffer, "word/comments.xml")).toContain(`w:date="${DATE}"`);
    expect(decode(buffer, "docProps/core.xml")).toContain(
      `<dcterms:created xsi:type="dcterms:W3CDTF">${DATE}`,
    );
    expect(decode(buffer, "docProps/core.xml")).toContain(
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${DATE}`,
    );
  });

  it("isolates concurrent generations from each other", async () => {
    const soloFirst = await scoped(DATE, () => generateDocument(options()));
    const soloSecond = await scoped(OTHER_DATE, () => generateDocument(options()));

    const [first, second] = await Promise.all([
      scoped(DATE, () => generateDocument(options())),
      scoped(OTHER_DATE, () => generateDocument(options())),
    ]);
    expect(first).toEqual(soloFirst);
    expect(second).toEqual(soloSecond);
    expect(first).not.toEqual(second);
  });

  it("keeps cryptographic ids and real dates outside a scope", () => {
    const first = generateDocumentSync(options());
    const second = generateDocumentSync(options());
    expect(first).not.toEqual(second);
    expect(decode(first, "docProps/core.xml")).not.toContain(DATE);
  });
});
