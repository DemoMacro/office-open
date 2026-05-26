import { Readable } from "stream";

import { afterEach, assert, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { createPacker, PrettifyType, type XmlifyedFile } from "./packer";

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

  describe("prettify passthrough", () => {
    it("should convert boolean true to WITH_2_BLANKS", async () => {
      await Packer.toString(mockFile, true);
      expect(compileMock).toHaveBeenCalledWith(
        expect.anything(),
        PrettifyType.WITH_2_BLANKS,
        expect.anything(),
      );
    });

    it("should pass PrettifyType value through", async () => {
      await Packer.toString(mockFile, PrettifyType.WITH_4_BLANKS);
      expect(compileMock).toHaveBeenCalledWith(
        expect.anything(),
        PrettifyType.WITH_4_BLANKS,
        expect.anything(),
      );
    });

    it("should convert boolean false to undefined", async () => {
      await Packer.toString(mockFile, false);
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), undefined, expect.anything());
    });

    it("should default overrides to empty array", async () => {
      await Packer.toString(mockFile);
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), undefined, []);
    });

    it("should pass overrides through", async () => {
      const overrides: readonly XmlifyedFile[] = [{ data: "test", path: "test.xml" }];
      await Packer.toString(mockFile, true, overrides);
      expect(compileMock).toHaveBeenCalledWith(expect.anything(), expect.anything(), overrides);
    });
  });

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
      const result = await Packer.pack(mockFile, "uint8array");
      expect(result).toBeInstanceOf(Uint8Array);
    });

    it("should export to binarystring", async () => {
      const result = await Packer.pack(mockFile, "binarystring");
      expect(typeof result).toBe("string");
    });

    it("should export to array", async () => {
      const result = await Packer.pack(mockFile, "array");
      expect(Array.isArray(result)).toBe(true);
    });
  });

  describe("#toStream()", () => {
    it("should return a Readable stream", () => {
      const stream = Packer.toStream(mockFile);
      expect(stream).toBeInstanceOf(Readable);
    });

    it("should emit ZIP magic bytes", async () => {
      const stream = Packer.toStream(mockFile);
      const chunks: Buffer[] = [];

      await new Promise<void>((resolve, reject) => {
        stream.on("data", (chunk: Buffer) => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", (err: Error) => reject(err));
      });

      const combined = Buffer.concat(chunks);
      expect(combined.length).toBeGreaterThan(0);
      // ZIP magic: PK\x03\x04
      expect(combined[0]).toBe(0x50);
      expect(combined[1]).toBe(0x4b);
    });

    it("should produce a valid ZIP with EOCD", async () => {
      const stream = Packer.toStream(mockFile);
      const chunks: Buffer[] = [];

      await new Promise<void>((resolve, reject) => {
        stream.on("data", (chunk: Buffer) => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", (err: Error) => reject(err));
      });

      const data = Buffer.concat(chunks);
      const eocd = data.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
      expect(eocd).toBeGreaterThanOrEqual(0);
    });

    it("should destroy the stream if compile throws", () => {
      compileMock.mockImplementation(() => {
        throw new Error("compile failed");
      });

      const stream = Packer.toStream(mockFile);

      return new Promise<void>((resolve) => {
        stream.on("error", (err: Error) => {
          expect(err.message).toBe("compile failed");
          resolve();
        });
        stream.on("end", () => {
          assert.fail("stream should not end normally on error");
        });
      });
    });

    it("should wrap non-Error throws in Error", () => {
      compileMock.mockImplementation(() => {
        throw "string error";
      });

      const stream = Packer.toStream(mockFile);

      return new Promise<void>((resolve) => {
        stream.on("error", (err: Error) => {
          expect(err).toBeInstanceOf(Error);
          expect(err.message).toBe("string error");
          resolve();
        });
        stream.on("end", () => {
          assert.fail("stream should not end normally on error");
        });
      });
    });

    it("should handle STORED entries (level 0)", async () => {
      const mockData = new TextEncoder().encode("test content");
      compileMock.mockReturnValue({
        "test.xml": [mockData, { level: 0 }],
      });

      const stream = Packer.toStream(mockFile);
      const chunks: Buffer[] = [];

      await new Promise<void>((resolve) => {
        stream.on("data", (chunk: Buffer) => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", () => resolve());
      });

      expect(Buffer.concat(chunks).length).toBeGreaterThan(0);
    });

    it("should handle DEFLATE entries with explicit level", async () => {
      const mockData = new TextEncoder().encode("test content");
      compileMock.mockReturnValue({
        "test.xml": [mockData, { level: 9 }],
      });

      const stream = Packer.toStream(mockFile);
      const chunks: Buffer[] = [];

      await new Promise<void>((resolve) => {
        stream.on("data", (chunk: Buffer) => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", () => resolve());
      });

      expect(Buffer.concat(chunks).length).toBeGreaterThan(0);
    });

    it("should support prettify option", async () => {
      const stream = Packer.toStream(mockFile, PrettifyType.WITH_TAB);
      const chunks: Buffer[] = [];

      await new Promise<void>((resolve, reject) => {
        stream.on("data", (chunk: Buffer) => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", (err: Error) => reject(err));
      });

      expect(chunks.length).toBeGreaterThan(0);
    });
  });
});
