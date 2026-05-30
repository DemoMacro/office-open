import type { FileChild } from "@file/file-child";
import { describe, expect, it } from "vite-plus/test";

import { StructuredDocumentTagBlock } from "./sdt";

describe("StructuredDocumentTagBlock", () => {
  it("should be assignable to FileChild (block-level element)", () => {
    const sdt: FileChild = new StructuredDocumentTagBlock({
      properties: { richText: true },
    });
    expect(sdt.fileChild).to.be.a("symbol");
  });
});
