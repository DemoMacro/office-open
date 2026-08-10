import { describe, expect, it } from "vite-plus/test";

import { Relationships, TargetModeType } from "./relationships";

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

  it("serialize() escapes special characters in target and id", () => {
    const rels = new Relationships();
    rels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      "a&b<c",
    );
    expect(rels.serialize()).toContain('Target="a&amp;b&lt;c"');
  });
});
