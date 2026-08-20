import { describe, expect, it } from "vitest";

import { zipSync } from "./packer";
import { parseArchive } from "./parser";
import { collectPassthroughParts } from "./passthrough";

function archiveOf(files: Record<string, string | Uint8Array>): ReturnType<typeof parseArchive> {
  const zippable: Record<string, Uint8Array> = {};
  for (const [path, data] of Object.entries(files)) {
    zippable[path] = typeof data === "string" ? new TextEncoder().encode(data) : data;
  }
  return parseArchive(zipSync(zippable as never));
}

const CONTENT_TYPES =
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleobject"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/customXml/item1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>` +
  `</Types>`;

const ROOT_RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

const DOC_RELS =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
  `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>` +
  `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
  `</Relationships>`;

const THEME_RELS =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/texture1.jpeg"/>` +
  `</Relationships>`;

function basePackage(): Record<string, string | Uint8Array> {
  return {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS,
    "word/document.xml": "<w:document/>",
    "word/_rels/document.xml.rels": DOC_RELS,
    "word/theme/theme1.xml": "<a:theme/>",
    "word/theme/_rels/theme1.xml.rels": THEME_RELS,
    "customXml/item1.xml": "<b:customXml/>",
    "customXml/itemProps1.xml": "<b:props/>",
    "word/media/image1.png": new Uint8Array([1, 2, 3]),
    "word/media/texture1.jpeg": new Uint8Array([4, 5, 6]),
    "word/embeddings/oleObject1.bin": new Uint8Array([7, 8]),
  };
}

describe("collectPassthroughParts", () => {
  it("keeps every part not listed as rebuilt, with companion rels and referenced-only media", () => {
    // document.xml AND its rels are rebuilt (relationships are renumbered);
    // everything else survives.
    const result = collectPassthroughParts(archiveOf(basePackage()), [
      "word/document.xml",
      "word/_rels/document.xml.rels",
    ]);
    const paths = result.parts.map((p) => p.path).sort();
    // theme + its rels + texture media (referenced only through the theme
    // rels), customXml pair, unabsorbed image, embedding — all survive.
    expect(paths).toEqual([
      "customXml/item1.xml",
      "customXml/itemProps1.xml",
      "word/embeddings/oleObject1.bin",
      "word/media/image1.png",
      "word/media/texture1.jpeg",
      "word/theme/_rels/theme1.xml.rels",
      "word/theme/theme1.xml",
    ]);
  });

  it("borrows Override content type first, then the extension Default", () => {
    const result = collectPassthroughParts(archiveOf(basePackage()), ["word/document.xml"]);
    const byPath = new Map(result.parts.map((p) => [p.path, p]));
    expect(byPath.get("customXml/item1.xml")?.contentType).toBe(
      "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
    );
    expect(byPath.get("word/embeddings/oleObject1.bin")?.contentType).toBe(
      "application/vnd.openxmlformats-officedocument.oleobject",
    );
  });

  it("omits contentType when the source declared neither Override nor Default", () => {
    const files = basePackage();
    files["mystery.part"] = new Uint8Array([9]);
    const result = collectPassthroughParts(archiveOf(files), ["word/document.xml"]);
    const mystery = result.parts.find((p) => p.path === "mystery.part");
    expect(mystery).toBeDefined();
    expect(mystery && "contentType" in mystery).toBe(false);
  });

  it("captures rebuilt→passthrough relationships, drops external targets", () => {
    const result = collectPassthroughParts(archiveOf(basePackage()), ["word/document.xml"]);
    const docRels = result.relationships.filter((r) => r.source === "word/document.xml");
    expect(
      docRels
        .map((r) => r.relationshipType.split("/").pop() ?? "")
        .sort((a, b) => a.localeCompare(b)),
    ).toEqual(["customXml", "image", "theme"]);
    // Targets stay in source-relative form.
    const theme = docRels.find((r) => r.relationshipType.endsWith("/theme"));
    expect(theme?.target).toBe("theme/theme1.xml");
    const customXml = docRels.find((r) => r.relationshipType.endsWith("/customXml"));
    expect(customXml?.target).toBe("../customXml/item1.xml");
  });

  it("does not capture relationships whose target is itself rebuilt", () => {
    // image1.png absorbed into the model this time — its relationship is the
    // compiler's own wiring, not a passthrough one.
    const result = collectPassthroughParts(archiveOf(basePackage()), [
      "word/document.xml",
      "word/media/image1.png",
      "word/_rels/document.xml.rels",
    ]);
    expect(result.parts.map((p) => p.path)).not.toContain("word/media/image1.png");
    const types = result.relationships.map((r) => r.relationshipType.split("/").pop());
    expect(types).not.toContain("image");
  });

  it("keeps a rebuilt part's rels verbatim when only the part is listed (glossary pattern)", () => {
    const files = basePackage();
    files["word/glossary/document.xml"] = "<w:glossary/>";
    files["word/glossary/_rels/document.xml.rels"] = ROOT_RELS;
    // glossary document rebuilt; its rels NOT listed → rels passes through.
    const result = collectPassthroughParts(archiveOf(files), ["word/glossary/document.xml"]);
    const paths = result.parts.map((p) => p.path);
    expect(paths).not.toContain("word/glossary/document.xml");
    expect(paths).toContain("word/glossary/_rels/document.xml.rels");
  });

  it("normalizes obsolete namespace URIs in passthrough XML without touching binaries", () => {
    const files = basePackage();
    files["word/stylesWithEffects.xml"] =
      '<?xml version="1.0"?>\r\n<w:styles xmlns:w15="http://schemas.microsoft.com/office/word/2010/11/wordml"><w:t>http://schemas.microsoft.com/office/word/2010/11/wordml</w:t></w:styles>';
    const binary = new TextEncoder().encode(
      "http://schemas.microsoft.com/office/word/2010/11/wordml",
    );
    files["word/media/namespace.bin"] = binary;

    const result = collectPassthroughParts(archiveOf(files), ["word/document.xml"]);
    const styles = result.parts.find((p) => p.path === "word/stylesWithEffects.xml");
    const media = result.parts.find((p) => p.path === "word/media/namespace.bin");

    expect(new TextDecoder().decode(styles?.data)).toBe(
      '<?xml version="1.0"?>\r\n<w:styles xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:t>http://schemas.microsoft.com/office/word/2010/11/wordml</w:t></w:styles>',
    );
    expect(Array.from(media?.data ?? [])).toEqual(Array.from(binary));
  });
});
