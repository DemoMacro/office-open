import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentationStream, generatePresentationSync } from "./generate";
import type { PresentationOptions } from "./shared/file";

const DATE = "2024-05-06T07:08:09.000Z";

const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** Shapes and pictures exercise the slide drawing-object id counters. */
const options = (): PresentationOptions => ({
  slides: [
    {
      children: [
        {
          shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "A" } },
        },
        {
          picture: { type: "png", x: 0, y: 0, width: 100, height: 100, data: PNG },
        },
      ],
    },
    {
      children: [{ shape: { x: 10, y: 10, width: 300, height: 150 } }],
    },
  ],
});

const decode = (buffer: Uint8Array, path: string): string =>
  new TextDecoder().decode(unzipSync(buffer)[path]);

describe("reproducible generation", () => {
  it("generates byte-identical presentations on repeated runs", () => {
    const first = generatePresentationSync(options(), { reproducible: { date: DATE } });
    const second = generatePresentationSync(options(), { reproducible: { date: DATE } });
    expect(Array.from(first)).toEqual(Array.from(second));
  });

  it("repeats byte-identically on the stream path", async () => {
    const collect = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> => {
      const reader = stream.getReader();
      const chunks: Uint8Array[] = [];
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
      }
      const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
      let offset = 0;
      for (const c of chunks) {
        out.set(c, offset);
        offset += c.length;
      }
      return out;
    };
    const stream = await collect(
      generatePresentationStream(options(), { reproducible: { date: DATE } }),
    );
    const sync = generatePresentationSync(options(), { reproducible: { date: DATE } });
    expect(Array.from(stream)).toEqual(Array.from(sync));
  });

  it("keeps slide drawing-object ids deterministic across runs", () => {
    const buffer = generatePresentationSync(options(), { reproducible: { date: DATE } });
    const slide1 = decode(buffer, "ppt/slides/slide1.xml");
    const ids = [...slide1.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]);
    // group shape → 1, shape → 2, picture → 3.
    expect(ids).toEqual(["1", "2", "3"]);
    const second = generatePresentationSync(options(), { reproducible: { date: DATE } });
    expect(decode(second, "ppt/slides/slide1.xml")).toBe(slide1);
  });

  it("drives the core-properties date default from the reproducible option", () => {
    const buffer = generatePresentationSync(options(), { reproducible: { date: DATE } });
    expect(decode(buffer, "docProps/core.xml")).toContain(DATE);
  });

  it("keeps non-deterministic ids without the reproducible option", () => {
    const first = generatePresentationSync(options());
    const second = generatePresentationSync(options());
    // Module-global counters keep incrementing, so run two differs from run one.
    expect(Array.from(first)).not.toEqual(Array.from(second));
  });
});
