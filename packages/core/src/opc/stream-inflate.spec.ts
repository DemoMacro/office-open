import { crc32 } from "node:zlib";

import { deflateSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import type { ByteSource } from "./byte-source";
import { inflateEntryFrom, STREAMING_XML_THRESHOLD } from "./stream-inflate";
import type { ZipEntryMeta } from "./zip-native";

function bufferSource(data: Uint8Array): ByteSource {
  return {
    byteLength: data.byteLength,
    async read(offset, length) {
      return data.slice(offset, offset + length);
    },
    stream(offset, length) {
      let done = false;
      return new ReadableStream<Uint8Array>({
        pull(controller) {
          if (done) return;
          controller.enqueue(data.slice(offset, offset + length));
          done = true;
          controller.close();
        },
      }) as ReadableStream<Uint8Array<ArrayBuffer>>;
    },
  };
}

describe("streaming OPC inflate", () => {
  it("verifies CRC before stripping oversized graphics data", async () => {
    const graphics = "A".repeat(1024 * 1024 + 1);
    const xml = new TextEncoder().encode(
      `<root><body>${"x".repeat(STREAMING_XML_THRESHOLD + 1)}</body>` +
        `<shape o:gfxdata="${graphics}"/></root>`,
    );
    const compressed = deflateSync(xml, { level: 6 });
    const entry: ZipEntryMeta = {
      name: "word/document.xml",
      method: 8,
      compSize: compressed.byteLength,
      uncompSize: xml.byteLength,
      crc: crc32(xml) >>> 0,
      dataStart: 0,
    };

    const result = await inflateEntryFrom(bufferSource(compressed), entry);

    expect(result.byteLength).toBeLessThan(xml.byteLength - graphics.length);
    expect(new TextDecoder().decode(result)).toContain("<shape/>");
  });

  it("rejects a mismatch on the original inflated stream", async () => {
    const xml = new TextEncoder().encode(`<root>${"x".repeat(STREAMING_XML_THRESHOLD + 1)}</root>`);
    const compressed = deflateSync(xml, { level: 6 });
    const entry: ZipEntryMeta = {
      name: "word/document.xml",
      method: 8,
      compSize: compressed.byteLength,
      uncompSize: xml.byteLength,
      crc: 1,
      dataStart: 0,
    };

    await expect(inflateEntryFrom(bufferSource(compressed), entry)).rejects.toThrow("CRC-32");
  });
});
