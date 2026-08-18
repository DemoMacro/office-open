import { describe, expect, it } from "vitest";

import { EmbeddingCollection } from "./embeddings";

describe("EmbeddingCollection", () => {
  it("allocates sequential names that skip pinned ones", () => {
    const c = new EmbeddingCollection();
    c.addEmbedding(new Uint8Array([1, 2]), "oleObject1.bin");
    expect(c.addEmbedding(new Uint8Array([3])).fileName).toBe("oleObject2.bin");
  });

  it("shares one entry for identical bytes under the same name", () => {
    const c = new EmbeddingCollection();
    const bytes = new Uint8Array([9, 9, 9]);
    const a = c.addEmbedding(bytes, "Microsoft_Excel_Worksheet1.xlsx");
    const b = c.addEmbedding(new Uint8Array([9, 9, 9]), "Microsoft_Excel_Worksheet1.xlsx");
    expect(b).toBe(a);
    expect(c.array).toHaveLength(1);
  });

  it("reallocates when a pinned name is claimed by different bytes", () => {
    const c = new EmbeddingCollection();
    const pinned = c.addEmbedding(new Uint8Array([1, 2, 3]), "oleObject1.bin");
    const fresh = c.addEmbedding(new Uint8Array([4, 5]), "oleObject1.bin");
    expect(fresh.fileName).not.toBe("oleObject1.bin");
    expect(c.array).toHaveLength(2);
    // The pinned payload survives untouched.
    expect(c.array.find((e) => e.fileName === pinned.fileName)?.data).toEqual(
      new Uint8Array([1, 2, 3]),
    );
  });
});
