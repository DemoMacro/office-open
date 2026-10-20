import { describe, expect, it } from "vite-plus/test";

import { auditCanonicalOptions } from "./raw-audit";

describe("canonical options rawParts audit", () => {
  it("allows orphaned independent members before format policy matching", () => {
    const blockers = auditCanonicalOptions(
      { rawParts: [{ path: "attachedToolbars.bin" }, { path: "embedded-package.docx" }] },
      [],
      new Set(["attachedToolbars.bin", "embedded-package.docx"]),
    );
    expect(blockers).toEqual([
      { part: "attachedToolbars.bin", reason: "orphaned-independent-part" },
      { part: "embedded-package.docx", reason: "orphaned-independent-part" },
    ]);
  });

  it("rejects canonical rawXml outside foreign-namespace policy", () => {
    const blockers = auditCanonicalOptions(
      { settings: { rawXml: "<w:unknown/>" }, rawParts: [{ path: "word/document.xml" }] },
      [],
      undefined,
      undefined,
      new Set(["word/document.xml"]),
    );
    expect(blockers).toEqual([
      { part: "settings.rawXml", reason: "canonical-options-rawXml" },
      { part: "word/document.xml", reason: "modeled-xml-passthrough" },
    ]);
  });
});
