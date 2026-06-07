import { afterEach, assert, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { createPacker, type XmlifyedFile } from "./packer";

// Simple mock compile function for testing createPacker
const compileMock = vi.fn();

const mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const Packer = createPacker<{ sections: readonly unknown[] }>({
  compile: compileMock,
  mimeType,
});

const mockFile = { sections: [] };

describe("createPacker", () => {
  beforeEach(() => {
    // Default: return minimal valid Zippable
    compileMock.mockReturnValue({ "test.xml": new TextEncoder().encode("<root/>") });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe("compile passthrough", () => {
    it("should default overrides to empty array", async () => {
      await Packer.toString(mockFile);
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), [], 0);
    });

    it("should pass overrides through", async () => {
      const overrides: readonly XmlifyedFile[] = [{ data: "test", path: "test.xml" }];
      await Packer.toString(mockFile, { overrides });
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), overrides, 0);
    });
  });

  describe("compression options", () => {
    it("should default to mediaLevel 0 (STORE) when no options", async () => {
      await Packer.toBuffer(mockFile);
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), [], 0);
    });

    it("should pass mediaLevel 0 when xml option is set", async () => {
      await Packer.toBuffer(mockFile, { compression: { xml: 9 } });
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), [], 0);
    });

    it("should pass custom mediaLevel", async () => {
      await Packer.toBuffer(mockFile, { compression: { media: 6 } });
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), [], 6);
    });

    it("should pass both xml and media options", async () => {
      await Packer.toBuffer(mockFile, { compression: { xml: 9, media: 4 } });
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), [], 4);
    });

    it("should work with sync methods", () => {
      Packer.toBufferSync(mockFile, { compression: { xml: 6, media: 0 } });
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), [], 0);
    });

    it("should produce valid output with custom compression", async () => {
      const result = await Packer.toBuffer(mockFile, { compression: { xml: 9 } });
      expect(result).toBeInstanceOf(Buffer);
      expect(result.byteLength).toBeGreaterThan(0);
    });

    it("should produce valid output with xml=0 (all STORE)", async () => {
      const result = await Packer.toBuffer(mockFile, { compression: { xml: 0 } });
      expect(result).toBeInstanceOf(Buffer);
      expect(result.byteLength).toBeGreaterThan(0);
    });
  });

  // ── Async methods ──

  describe("#toString()", () => {
    it("should return a non-empty string", async () => {
      const result = await Packer.toString(mockFile);
      expect(result.length).toBeGreaterThan(0);
    });
  });

  describe("#toBuffer()", () => {
    it("should return a Buffer", async () => {
      const result = await Packer.toBuffer(mockFile);
      expect(result).toBeInstanceOf(Buffer);
      expect(result.byteLength).toBeGreaterThan(0);
    });

    it("should reject when compile throws", async () => {
      compileMock.mockImplementation(() => {
        throw new Error("compile failed");
      });
      await expect(Packer.toBuffer(mockFile)).rejects.toBeDefined();
    });
  });

  describe("#toBytes()", () => {
    it("should return a Uint8Array", async () => {
      const result = await Packer.toBytes(mockFile);
      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.byteLength).toBeGreaterThan(0);
    });
  });

  describe("#toBase64String()", () => {
    it("should return a base64 string", async () => {
      const result = await Packer.toBase64String(mockFile);
      expect(result).toBeDefined();
      expect(result.length).toBeGreaterThan(0);
    });
  });

  describe("#toBlob()", () => {
    it("should return a Blob with correct MIME type", async () => {
      const result = await Packer.toBlob(mockFile);
      expect(result).toBeInstanceOf(Blob);
      expect(result.type).toBe(mimeType);
    });
  });

  describe("#toArrayBuffer()", () => {
    it("should return an ArrayBuffer", async () => {
      const result = await Packer.toArrayBuffer(mockFile);
      expect(result).toBeInstanceOf(ArrayBuffer);
      expect(result.byteLength).toBeGreaterThan(0);
    });
  });

  describe("#pack()", () => {
    it("should export to uint8array", async () => {
      const result = await Packer.pack(mockFile, { type: "uint8array" });
      expect(result).toBeInstanceOf(Uint8Array);
    });

    it("should export to binarystring", async () => {
      const result = await Packer.pack(mockFile, { type: "binarystring" });
      expect(typeof result).toBe("string");
    });

    it("should export to array", async () => {
      const result = await Packer.pack(mockFile, { type: "array" });
      expect(Array.isArray(result)).toBe(true);
    });
  });

  // ── Sync methods ──

  describe("#toBufferSync()", () => {
    it("should return a Buffer synchronously", () => {
      const result = Packer.toBufferSync(mockFile);
      expect(result).toBeInstanceOf(Buffer);
      expect(result.byteLength).toBeGreaterThan(0);
    });

    it("should throw when compile throws", () => {
      compileMock.mockImplementation(() => {
        throw new Error("compile failed");
      });
      expect(() => Packer.toBufferSync(mockFile)).toThrow("compile failed");
    });
  });

  describe("#toBytesSync()", () => {
    it("should return a Uint8Array synchronously", () => {
      const result = Packer.toBytesSync(mockFile);
      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.byteLength).toBeGreaterThan(0);
    });
  });

  describe("#toStringSync()", () => {
    it("should return a string synchronously", () => {
      const result = Packer.toStringSync(mockFile);
      expect(typeof result).toBe("string");
      expect(result.length).toBeGreaterThan(0);
    });
  });

  describe("#toBase64StringSync()", () => {
    it("should return a base64 string synchronously", () => {
      const result = Packer.toBase64StringSync(mockFile);
      expect(result.length).toBeGreaterThan(0);
    });
  });

  describe("#toBlobSync()", () => {
    it("should return a Blob synchronously", () => {
      const result = Packer.toBlobSync(mockFile);
      expect(result).toBeInstanceOf(Blob);
      expect(result.type).toBe(mimeType);
    });
  });

  describe("#toArrayBufferSync()", () => {
    it("should return an ArrayBuffer synchronously", () => {
      const result = Packer.toArrayBufferSync(mockFile);
      expect(result).toBeInstanceOf(ArrayBuffer);
      expect(result.byteLength).toBeGreaterThan(0);
    });
  });

  describe("#packSync()", () => {
    it("should export to uint8array synchronously", () => {
      const result = Packer.packSync(mockFile, { type: "uint8array" });
      expect(result).toBeInstanceOf(Uint8Array);
    });

    it("should export to nodebuffer synchronously", () => {
      const result = Packer.packSync(mockFile, { type: "nodebuffer" });
      expect(result).toBeInstanceOf(Buffer);
    });
  });

  // ── Async/Sync parity ──

  describe("async/sync parity", () => {
    it("should produce identical output for toBuffer vs toBufferSync", async () => {
      const asyncResult = await Packer.toBuffer(mockFile);
      const syncResult = Packer.toBufferSync(mockFile);
      expect(asyncResult.equals(syncResult)).toBe(true);
    });

    it("should produce identical output for toBytes vs toBytesSync", async () => {
      const asyncResult = await Packer.toBytes(mockFile);
      const syncResult = Packer.toBytesSync(mockFile);
      expect(asyncResult).toEqual(syncResult);
    });
  });

  // ── Stream ──

  /** Helper: collect all chunks from a ReadableStream<Uint8Array>. */
  const collectStream = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> => {
    const chunks: Uint8Array[] = [];
    const reader = stream.getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total = chunks.reduce((s, c) => s + c.length, 0);
    const combined = new Uint8Array(total);
    let offset = 0;
    for (const c of chunks) {
      combined.set(c, offset);
      offset += c.length;
    }
    return combined;
  };

  describe("#toStream()", () => {
    it("should return a ReadableStream", () => {
      const stream = Packer.toStream(mockFile);
      expect(stream).toBeInstanceOf(ReadableStream);
    });

    it("should emit ZIP magic bytes", async () => {
      const stream = Packer.toStream(mockFile);
      const combined = await collectStream(stream);
      expect(combined.length).toBeGreaterThan(0);
      // ZIP magic: PK\x03\x04
      expect(combined[0]).toBe(0x50);
      expect(combined[1]).toBe(0x4b);
    });

    it("should produce a valid ZIP with EOCD", async () => {
      const stream = Packer.toStream(mockFile);
      const data = await collectStream(stream);
      const eocdSignature = new Uint8Array([0x50, 0x4b, 0x05, 0x06]);
      let found = false;
      for (let i = 0; i <= data.length - 4; i++) {
        if (
          data[i] === eocdSignature[0] &&
          data[i + 1] === eocdSignature[1] &&
          data[i + 2] === eocdSignature[2] &&
          data[i + 3] === eocdSignature[3]
        ) {
          found = true;
          break;
        }
      }
      expect(found).toBe(true);
    });

    it("should error the stream if compile throws", async () => {
      compileMock.mockImplementation(() => {
        throw new Error("compile failed");
      });

      const stream = Packer.toStream(mockFile);
      const reader = stream.getReader();

      await expect(reader.read()).rejects.toThrow("compile failed");
    });

    it("should wrap non-Error throws in Error", async () => {
      compileMock.mockImplementation(() => {
        throw "string error";
      });

      const stream = Packer.toStream(mockFile);
      const reader = stream.getReader();

      try {
        await reader.read();
        assert.fail("should have thrown");
      } catch (err) {
        expect(err).toBeInstanceOf(Error);
        expect((err as Error).message).toBe("string error");
      }
    });

    it("should handle STORED entries (level 0)", async () => {
      const mockData = new TextEncoder().encode("test content");
      compileMock.mockReturnValue({
        "test.xml": [mockData, { level: 0 }],
      });

      const stream = Packer.toStream(mockFile);
      const data = await collectStream(stream);
      expect(data.length).toBeGreaterThan(0);
    });

    it("should handle DEFLATE entries with explicit level", async () => {
      const mockData = new TextEncoder().encode("test content");
      compileMock.mockReturnValue({
        "test.xml": [mockData, { level: 9 }],
      });

      const stream = Packer.toStream(mockFile);
      const data = await collectStream(stream);
      expect(data.length).toBeGreaterThan(0);
    });
  });
});
