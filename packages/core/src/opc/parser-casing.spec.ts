import { describe, expect, it } from "vitest";

import { zipSync } from "./packer";
import { parseArchive } from "./parser";

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

describe("ParsedArchive case-insensitive part names", () => {
  const archiveOf = (): ReturnType<typeof parseArchive> =>
    parseArchive(
      zipSync({
        "_rels/.rels": [
          enc(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
          ),
          { level: 0 },
        ],
        "xl/SharedStrings.xml": [enc("<sst/>"), { level: 0 }],
        "xl/media/image1.PNG": [new Uint8Array([1, 2, 3]), { level: 0 }],
      }),
    );

  it("resolves rel targets spelled with different casing than the entry", () => {
    const doc = archiveOf();
    expect(doc.get("xl/sharedStrings.xml")?.name).toBe("sst");
    expect(doc.getRaw("xl/media/image1.png")).toBeDefined();
    expect(doc.has("xl/sharedstrings.xml")).toBe(true);
  });

  it("resolvePath returns the canonical stored path", () => {
    const doc = archiveOf();
    expect(doc.resolvePath("xl/sharedstrings.xml")).toBe("xl/SharedStrings.xml");
    expect(doc.resolvePath("xl/missing.xml")).toBe("xl/missing.xml");
    expect(doc.keys()).toContain("xl/SharedStrings.xml");
  });

  it("remove clears the case-insensitive mapping", () => {
    const doc = archiveOf();
    expect(doc.remove("xl/sharedstrings.xml")).toBe(true);
    expect(doc.has("xl/SharedStrings.xml")).toBe(false);
  });
});
