import { describe, expect, it } from "vite-plus/test";

import type { Zippable } from "../opc/packer";
import { addBinaryFile, compileMapping } from "./compile";

const decode = (value: unknown): string => new TextDecoder().decode(value as Uint8Array);

function tupleLevel(value: unknown): number | undefined {
  return (value as [Uint8Array, { level: number }])[1].level;
}

describe("compileMapping", () => {
  it("flattens optional arrays and preserves binary data", () => {
    const binary = new Uint8Array([1, 2, 3]);
    const files = compileMapping({
      document: { path: "word/document.xml", data: "<document/>" },
      headers: [
        { path: "word/header1.xml", data: "<header/>" },
        { path: "word/header2.bin", data: binary },
      ],
      absent: undefined,
    });

    expect(decode(files["word/document.xml"])).toBe("<document/>");
    expect(decode(files["word/header1.xml"])).toBe("<header/>");
    expect(files["word/header2.bin"]).toBe(binary);
  });

  it("applies overrides after mapping entries", () => {
    const files = compileMapping({ document: { path: "word/document.xml", data: "before" } }, [
      { path: "word/document.xml", data: "after" },
    ]);
    expect(decode(files["word/document.xml"])).toBe("after");
  });
});

describe("addBinaryFile", () => {
  it("stores precompressed media and applies media level to compressible data", () => {
    const files: Zippable = {};
    addBinaryFile(files, "ppt/media/image1.png", new Uint8Array([1]), 6);
    addBinaryFile(files, "ppt/media/image2.svg", new Uint8Array([2]), 6);

    expect(tupleLevel(files["ppt/media/image1.png"])).toBe(0);
    expect(tupleLevel(files["ppt/media/image2.svg"])).toBe(6);
  });
});
