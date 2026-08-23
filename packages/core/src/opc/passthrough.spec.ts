import { describe, expect, it } from "vitest";

import { zipSync } from "./packer";
import { parseArchive } from "./parser";
import { collectPassthroughParts, dropDanglingPassthroughRels } from "./passthrough";

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

const ROOT_RELS =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>` +
  `<Relationship Id="rId3" Type="urn:example:external-root" Target="https://example.com/package" TargetMode="External"/>` +
  `</Relationships>`;

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
    "docProps/thumbnail.jpeg": new Uint8Array([0xff, 0xd8, 0xff]),
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
      "docProps/thumbnail.jpeg",
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

  it("captures internal passthrough targets and drops rebuilt-owner external targets", () => {
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
    expect(docRels.some((r) => r.targetMode === "External")).toBe(false);
  });

  it("captures internal and external package-root relationships", () => {
    const result = collectPassthroughParts(archiveOf(basePackage()), ["word/document.xml"]);
    const rootRels = result.relationships.filter((r) => r.source === "");
    expect(rootRels).toEqual([
      {
        source: "",
        relationshipType:
          "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
        target: "docProps/thumbnail.jpeg",
        rId: "rId2",
      },
      {
        source: "",
        relationshipType: "urn:example:external-root",
        target: "https://example.com/package",
        rId: "rId3",
        targetMode: "External",
      },
    ]);
  });

  it("matches passthrough targets case-insensitively (OPC part-name rule)", () => {
    // A source can spell the rel target with different casing than the ZIP
    // entry; Office still resolves it, so the relationship must survive.
    const files = basePackage();
    delete files["docProps/thumbnail.jpeg"];
    files["docProps/Thumbnail.jpeg"] = new Uint8Array([0xff, 0xd8, 0xff]);
    const result = collectPassthroughParts(archiveOf(files), ["word/document.xml"]);
    const rootRels = result.relationships.filter((r) => r.source === "");
    expect(rootRels.map((r) => r.relationshipType.split("/").pop())).toContain("thumbnail");
  });

  it("keeps a rebuilt-target relationship captured for source-id pre-claim", () => {
    // image1.png absorbed into the model this time — the part no longer passes
    // through, but its source rel stays captured: the compiler's pre-claim
    // reads the source rId to re-emit the rel at the same id, so verbatim
    // body content referencing it survives round-trip.
    const result = collectPassthroughParts(archiveOf(basePackage()), [
      "word/document.xml",
      "word/media/image1.png",
      "word/_rels/document.xml.rels",
    ]);
    expect(result.parts.map((p) => p.path)).not.toContain("word/media/image1.png");
    const image = result.relationships.find((r) => r.relationshipType.split("/").pop() === "image");
    expect(image?.target).toBe("media/image1.png");
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

  it("normalizes obsolete URIs in single-quoted and long-prefixed declarations", () => {
    const files = basePackage();
    files["word/odd.xml"] =
      "<r xmlns:wordprocessingCanvas='http://schemas.openxmlformats.org/wordprocessingml/2006/3/main'><a/></r>";

    const result = collectPassthroughParts(archiveOf(files), ["word/document.xml"]);
    const odd = result.parts.find((p) => p.path === "word/odd.xml");
    expect(new TextDecoder().decode(odd?.data)).toBe(
      "<r xmlns:wordprocessingCanvas='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><a/></r>",
    );
  });

  it("leaves obsolete URIs in non-xmlns attribute values untouched", () => {
    const files = basePackage();
    files["word/attr.xml"] =
      '<d href="http://schemas.openxmlformats.org/spreadsheetml/2006/5/main"/>';
    const result = collectPassthroughParts(archiveOf(files), ["word/document.xml"]);
    const attr = result.parts.find((p) => p.path === "word/attr.xml");
    expect(new TextDecoder().decode(attr?.data)).toContain(
      'href="http://schemas.openxmlformats.org/spreadsheetml/2006/5/main"',
    );
  });
});

describe("dropDanglingPassthroughRels", () => {
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();

  /** A rebuilt package shape: document.xml (rebuilt) + rels + a theme part. */
  function assembled(): Record<string, Uint8Array> {
    return {
      "word/document.xml": encoder.encode("<w:document/>"),
      "word/_rels/document.xml.rels": encoder.encode(
        `<Relationships>` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
          `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
          `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>` +
          `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
          `</Relationships>`,
      ),
      "word/theme/theme1.xml": encoder.encode("<a:theme/>"),
      "word/media/image1.png": new Uint8Array([1]),
      "_rels/.rels": encoder.encode(
        `<Relationships>` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
          `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>` +
          `</Relationships>`,
      ),
    };
  }

  it("drops internal rels whose target part is absent", () => {
    const files = assembled();
    // theme (present) and customXml (absent — hand-authored without rawParts)
    const dropped = dropDanglingPassthroughRels(files, [
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
        target: "../customXml/item1.xml",
        rId: "rId2",
      },
    ]);
    expect(dropped).toBe(1);
    const rels = decoder.decode(files["word/_rels/document.xml.rels"]);
    expect(rels).not.toContain("customXml");
    // siblings survive
    expect(rels).toContain('Target="theme/theme1.xml"');
  });

  it("keeps rels whose target part exists", () => {
    const files = assembled();
    const dropped = dropDanglingPassthroughRels(files, [
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/theme1.xml",
        rId: "rId1",
      },
    ]);
    expect(dropped).toBe(0);
    expect(decoder.decode(files["word/_rels/document.xml.rels"])).toContain(
      'Target="theme/theme1.xml"',
    );
  });

  it("keeps External and #fragment targets without touching their rels part", () => {
    const files = assembled();
    const before = files["word/_rels/document.xml.rels"];
    const dropped = dropDanglingPassthroughRels(files, [
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.com",
        rId: "rId3",
        targetMode: "External",
      },
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "#Sheet2!A1",
        rId: "rId5",
      },
    ]);
    expect(dropped).toBe(0);
    expect(files["word/_rels/document.xml.rels"]).toBe(before);
  });

  it('drops root-level rels from _rels/.rels (source "")', () => {
    const files = assembled();
    const dropped = dropDanglingPassthroughRels(files, [
      {
        source: "",
        relationshipType:
          "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
        target: "docProps/thumbnail.jpeg",
        rId: "rId2",
      },
    ]);
    expect(dropped).toBe(1);
    const root = decoder.decode(files["_rels/.rels"]);
    expect(root).not.toContain("thumbnail");
    expect(root).toContain("officeDocument");
  });

  it("resolves targets case-insensitively (OPC part-name matching)", () => {
    const files = assembled();
    // part stored as Thumbnail.jpeg, rel target spelled thumbnail.jpeg
    delete files["word/media/image1.png"];
    files["word/Media/Image1.PNG"] = new Uint8Array([1]);
    const dropped = dropDanglingPassthroughRels(files, [
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: "media/image1.png",
        rId: "rId4",
      },
    ]);
    expect(dropped).toBe(0);
    expect(decoder.decode(files["word/_rels/document.xml.rels"])).toContain(
      'Target="media/image1.png"',
    );
  });

  it("preserves the [data, opts] entry form when rewriting", () => {
    const files = assembled();
    const opts = { level: 0 as const };
    files["word/_rels/document.xml.rels"] = encoder.encode(
      `<Relationships>` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
        `</Relationships>`,
    );
    // @ts-expect-error tuple entry — Zippable allows it
    files["word/_rels/document.xml.rels"] = [files["word/_rels/document.xml.rels"], opts];
    dropDanglingPassthroughRels(files, [
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
        target: "../customXml/item1.xml",
        rId: "rId1",
      },
    ]);
    expect(Array.isArray(files["word/_rels/document.xml.rels"])).toBe(true);
  });
});
