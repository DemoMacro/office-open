import { decodeBase64, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generate } from "../generate";
import { generateVerifiedBase64 } from "./verify";

const encode = (s: string) => new TextEncoder().encode(s);

describe("generateVerifiedBase64", () => {
  it("returns base64 for a clean freshly generated docx", async () => {
    const bytes = (await generate({
      type: "docx",
      options: { sections: [{ children: [{ paragraph: { children: ["Hello"] } }] }] },
      outputType: "uint8array",
    })) as Uint8Array;
    const base64 = generateVerifiedBase64("docx", bytes);
    // Copy into a plain Uint8Array — decodeBase64 may return a Buffer subclass
    // that vitest refuses to deep-equal against a plain Uint8Array.
    expect(new Uint8Array(decodeBase64(base64))).toEqual(bytes);
  });

  it("throws with the issue details when the package is inconsistent", () => {
    // Minimal broken package: duplicate relationship ids in one rels part.
    const broken = zipSync({
      "[Content_Types].xml": encode(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`,
      ),
      "_rels/.rels": encode(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="a" Target="x.xml"/>` +
          `<Relationship Id="rId1" Type="b" Target="y.xml"/></Relationships>`,
      ),
    });
    expect(() => generateVerifiedBase64("docx", broken)).toThrow(/O7 .*rId1/);
    expect(() => generateVerifiedBase64("docx", broken)).toThrow(/office-open bug/);
  });
});
