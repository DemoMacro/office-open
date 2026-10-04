import { describe, expect, it } from "vite-plus/test";

import { generateDocumentSync } from "../generate";
import { parseDocumentSync } from "../parse";
import type { DocumentOptions } from "./core-properties";

describe("comment metadata parts", () => {
  it("round-trips through a DOCX package", async () => {
    const options: DocumentOptions = {
      sections: [{ children: [] }],
      commentsIds: [{ paraId: "7FD6C115", durableId: "1EC7B1B1" }],
      commentsExtensible: [{ durableId: "1EC7B1B1", intelligentPlaceholder: true }],
    };
    const bytes = generateDocumentSync(options, { type: "uint8array" });
    const parsed = parseDocumentSync(bytes);
    expect(parsed.commentsIds).toEqual(options.commentsIds);
    expect(parsed.commentsExtensible).toEqual(options.commentsExtensible);
  });
});
