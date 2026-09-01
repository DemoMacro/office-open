import { describe, expect, it } from "vite-plus/test";

import { buildRootRelationships, Relationships, TargetModeType } from "./relationships";

describe("Relationships", () => {
  it("add() allocates sequential ids and returns the numeric id", () => {
    const rels = new Relationships();
    expect(
      rels.add(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        "theme/theme1.xml",
      ),
    ).toBe(1);
    expect(
      rels.add(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        "slides/slide1.xml",
      ),
    ).toBe(2);
    expect(rels.relationshipCount).toBe(2);
  });

  it("add() forwards targetMode for external relationships", () => {
    const rels = new Relationships();
    rels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      "https://example.com",
      TargetModeType.EXTERNAL,
    );
    expect(rels.serialize()).toContain('TargetMode="External"');
  });

  it("addRelationship() takes an explicit id (fixed rId1, offset batches)", () => {
    const rels = new Relationships();
    rels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      "theme/theme1.xml",
    );
    expect(rels.serialize()).toContain('Id="rId1"');
  });

  it("buildRootRelationships creates fixed package relationships with optional custom props", () => {
    const withoutCustom = buildRootRelationships("ppt/presentation.xml", false).serialize();
    expect(withoutCustom).toContain('Id="rId1"');
    expect(withoutCustom).toContain('Target="ppt/presentation.xml"');
    expect(withoutCustom).toContain('Target="docProps/core.xml"');
    expect(withoutCustom).toContain('Target="docProps/app.xml"');
    expect(withoutCustom).not.toContain("custom-properties");

    const withCustom = buildRootRelationships("word/document.xml", true).serialize();
    expect(withCustom).toContain('Id="rId4"');
    expect(withCustom).toContain('Target="docProps/custom.xml"');
  });

  it("buildRootRelationships preserves package-root external target mode", () => {
    const xml = buildRootRelationships("word/document.xml", false, [
      {
        source: "",
        relationshipType: "urn:example:external-root",
        target: "https://example.com/package",
        targetMode: "External",
      },
    ]).serialize();
    expect(xml).toContain(
      'Type="urn:example:external-root" Target="https://example.com/package" TargetMode="External"',
    );
  });

  it("serialize() escapes special characters in target and id", () => {
    const rels = new Relationships();
    rels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      "a&b<c",
    );
    expect(rels.serialize()).toContain('Target="a&amp;b&lt;c"');
  });

  it("reserveId() lifts the next-id watermark without registering an entry", () => {
    const rels = new Relationships();
    // Passthrough rIds arrive in the prefixed form ("rId7")
    rels.reserveId("rId7");
    expect(rels.relationshipCount).toBe(0);
    expect(rels.nextRelationshipId).toBe(8);
    // Numbers and bare repeated reserves stay idempotent
    rels.reserveId(3);
    rels.reserveId("rId7");
    expect(rels.nextRelationshipId).toBe(8);
    // A later source re-use can still claim the reserved slot
    rels.addRelationship(
      7,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
      "fontTable.xml",
    );
    expect(rels.nextRelationshipId).toBe(8);
  });

  it("reserveId() ignores malformed ids instead of corrupting the watermark", () => {
    const rels = new Relationships();
    rels.reserveId("rIdrId7");
    rels.reserveId("not-an-id");
    expect(rels.nextRelationshipId).toBe(1);
  });

  it("reserveSourceRids() pre-claims one source part's ids from the passthrough list", () => {
    const rels = new Relationships();
    rels.reserveSourceRids("word/document.xml", [
      { source: "word/document.xml", rId: "rId6" },
      { source: "word/document.xml", rId: "rId7" },
      // Another part's rels — must not lift this instance's watermark.
      { source: "word/styles.xml", rId: "rId5" },
    ]);
    expect(rels.relationshipCount).toBe(0);
    expect(rels.nextRelationshipId).toBe(8);
    // Auto-allocation lands above the source id space.
    expect(
      rels.add(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "comments.xml",
      ),
    ).toBe(8);
  });
});
