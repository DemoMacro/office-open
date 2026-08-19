import { describe, expect, it } from "vitest";

import {
  encryptedContainerOutput,
  encryptedContainerStream,
  isEncryptedContainer,
} from "./encryption";
import { OoxmlMimeType } from "./output";

const CFB_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00]);

describe("isEncryptedContainer", () => {
  it("matches the OLE2/CFB signature", () => {
    expect(isEncryptedContainer(CFB_BYTES)).toBe(true);
  });

  it("rejects ZIP packages and truncated inputs", () => {
    expect(isEncryptedContainer(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false);
    expect(isEncryptedContainer(CFB_BYTES.subarray(0, 7))).toBe(false);
    expect(isEncryptedContainer(new Uint8Array())).toBe(false);
  });
});

describe("encryptedContainerOutput", () => {
  it("converts the verbatim payload to the requested output type", async () => {
    const out = encryptedContainerOutput(
      { encrypted: { data: CFB_BYTES } },
      "uint8array",
      OoxmlMimeType.DOCX,
    );
    expect(out).toBe(CFB_BYTES);

    const base64 = encryptedContainerOutput(
      { encrypted: { data: Array.from(CFB_BYTES) } },
      "base64",
      OoxmlMimeType.XLSX,
    );
    expect(typeof base64).toBe("string");
  });

  it("returns undefined without an encrypted payload", () => {
    expect(encryptedContainerOutput({}, "uint8array", OoxmlMimeType.DOCX)).toBeUndefined();
  });
});

describe("encryptedContainerStream", () => {
  it("streams the payload unchanged", async () => {
    const stream = encryptedContainerStream({ encrypted: { data: CFB_BYTES } });
    expect(stream).toBeDefined();
    const chunks: Uint8Array[] = [];
    const reader = stream!.getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total = chunks.reduce((n, c) => n + c.length, 0);
    const bytes = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      bytes.set(c, off);
      off += c.length;
    }
    expect(bytes).toEqual(CFB_BYTES);
    expect(encryptedContainerStream({})).toBeUndefined();
  });
});
