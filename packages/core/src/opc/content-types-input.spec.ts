import { describe, expect, it } from "vite-plus/test";

import type { ContentTypesInput } from "./content-types-input";
import {
  contentTypesDesc,
  deriveContentTypes,
  finalizeContentTypes,
  resolverFromRegistry,
} from "./content-types-input";
import { PPTX_PARTS } from "./part-registry";

const resolve = resolverFromRegistry(PPTX_PARTS);
const MEDIA = {
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  emf: "image/x-emf",
  mp4: "video/mp4",
};

describe("resolverFromRegistry", () => {
  it("matches dense sequential and sparse index-based part names equally", () => {
    // Dense: slides are written slide1, slide2, slide3 in sequence.
    expect(resolve("ppt/slides/slide1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    expect(resolve("ppt/slides/slide3.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    // Sparse: comments are keyed by slide index, so comment1 and comment3 can
    // both exist with comment2 absent. The resolver must match either.
    expect(resolve("ppt/comments/comment1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
    );
    expect(resolve("ppt/comments/comment3.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
    );
  });

  it("returns undefined for media, embeddings, and .rels (Default-resolved)", () => {
    expect(resolve("ppt/media/image1.png")).toBeUndefined();
    expect(resolve("ppt/embeddings/oleObject1.bin")).toBeUndefined();
    expect(resolve("ppt/_rels/slide1.xml.rels")).toBeUndefined();
    expect(resolve("_rels/.rels")).toBeUndefined();
  });

  it("does not confuse a rels path with its part", () => {
    // The slide rels file lives next to the slide part; the slide pattern must
    // not match the rels path.
    expect(resolve("ppt/slides/_rels/slide1.xml.rels")).toBeUndefined();
  });
});

describe("deriveContentTypes", () => {
  it("emits Overrides for resolved parts (leading slash) and Defaults for media", () => {
    const input = deriveContentTypes(
      [
        "ppt/presentation.xml",
        "ppt/slides/slide1.xml",
        "ppt/comments/comment3.xml",
        "ppt/media/image1.png",
        "ppt/media/image2.png",
        "docProps/core.xml",
      ],
      { resolve, mediaContentTypes: MEDIA },
    );
    const partNames = input.overrides.map((o) => o.partName).sort();
    expect(partNames).toEqual(
      [
        "/docProps/core.xml",
        "/ppt/comments/comment3.xml",
        "/ppt/presentation.xml",
        "/ppt/slides/slide1.xml",
      ].sort(),
    );
    // png declared once despite two png files.
    expect(input.defaults.map((d) => d.extension)).toContain("png");
    expect(input.defaults.filter((d) => d.extension === "png")).toHaveLength(1);
  });

  it("always carries rels + xml base defaults, and does not duplicate them", () => {
    const input = deriveContentTypes(
      ["_rels/.rels", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"],
      { resolve, mediaContentTypes: MEDIA },
    );
    const exts = input.defaults.map((d) => d.extension);
    expect(exts).toContain("rels");
    expect(exts).toContain("xml");
    expect(exts.filter((e) => e === "rels")).toHaveLength(1);
  });

  it("skips [Content_Types].xml itself", () => {
    const input = deriveContentTypes(["[Content_Types].xml", "ppt/presentation.xml"], {
      resolve,
      mediaContentTypes: MEDIA,
    });
    expect(input.overrides).toHaveLength(1);
    expect(input.overrides[0]!.partName).toBe("/ppt/presentation.xml");
  });

  it("dedups extensions case-insensitively", () => {
    const input = deriveContentTypes(["ppt/media/A.PNG", "ppt/media/b.png"], {
      resolve,
      mediaContentTypes: MEDIA,
    });
    expect(input.defaults.filter((d) => d.extension.toLowerCase() === "png")).toHaveLength(1);
  });

  it("serializes derived input via contentTypesDesc", () => {
    const input: ContentTypesInput = deriveContentTypes(
      ["ppt/presentation.xml", "ppt/media/image1.png", "ppt/slides/slide1.xml"],
      { resolve, mediaContentTypes: MEDIA },
    );
    const xml = contentTypesDesc.stringify(input, {} as never)!;
    expect(xml).toContain('<Override PartName="/ppt/presentation.xml"');
    expect(xml).toContain('<Default Extension="png"');
    expect(xml).toContain('<Default Extension="rels"');
  });

  it("merges explicit overrides for data-driven parts (altChunks, sub-documents)", () => {
    // altChunk paths are not in the registry and carry a per-file content type.
    const input = deriveContentTypes(["word/afchunks/chunk1.html", "word/document.xml"], {
      resolve,
      mediaContentTypes: MEDIA,
      overrides: [{ path: "word/afchunks/chunk1.html", contentType: "text/html" }],
    });
    const partNames = input.overrides.map((o) => o.partName);
    expect(partNames).toContain("/word/afchunks/chunk1.html");
    expect(
      input.overrides.find((o) => o.partName === "/word/afchunks/chunk1.html")?.contentType,
    ).toBe("text/html");
  });

  it("lets an explicit override take precedence over a resolver match for the same path", () => {
    const input = deriveContentTypes(["ppt/presentation.xml"], {
      resolve,
      mediaContentTypes: MEDIA,
      overrides: [{ path: "ppt/presentation.xml", contentType: "application/custom+xml" }],
    });
    expect(input.overrides).toHaveLength(1);
    expect(input.overrides[0]!.contentType).toBe("application/custom+xml");
  });
});

describe("finalizeContentTypes", () => {
  it("preserves source declarations for absent parts and adds only uncovered outputs", () => {
    const source: ContentTypesInput = {
      preserveSourceDeclarations: true,
      defaults: [
        {
          extension: "rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
      ],
      overrides: [
        { partName: "/xl/styles.xml", contentType: "spreadsheetml.styles+xml" },
        { partName: "/docProps/core.xml", contentType: "core-properties+xml" },
      ],
    };
    const xml = finalizeContentTypes(
      { "xl/workbook.xml": new TextEncoder().encode("<workbook/>") },
      {
        resolve: (path) => (path === "xl/workbook.xml" ? "spreadsheetml.workbook+xml" : undefined),
        mediaContentTypes: {},
        source,
      },
      {} as never,
    );

    expect(xml).toContain('PartName="/docProps/core.xml"');
    expect(xml).toContain('Extension="rels"');
    expect(xml).toContain('PartName="/xl/workbook.xml"');
  });

  it("does not force an override when the source default already declares the same type", () => {
    const input = finalizeContentTypes(
      { "xl/workbook.xml": new TextEncoder().encode("<workbook/>") },
      {
        resolve: () => "application/vnd.custom.workbook+xml",
        mediaContentTypes: {},
        source: {
          defaults: [{ extension: "xml", contentType: "application/vnd.custom.workbook+xml" }],
          overrides: [],
        },
        forcedOverrides: [
          { path: "xl/workbook.xml", contentType: "application/vnd.custom.workbook+xml" },
        ],
      },
      {} as never,
    );
    expect(input).not.toContain("<Override");
  });

  it("keeps a generic source XML default instead of retyping the part", () => {
    const input = finalizeContentTypes(
      { "ppt/diagrams/drawing1.xml": new TextEncoder().encode("<drawing/>") },
      {
        resolve: () => "application/vnd.ms-office.drawingml.diagramDrawing+xml",
        mediaContentTypes: {},
        source: {
          preserveSourceDeclarations: true,
          defaults: [{ extension: "xml", contentType: "application/xml" }],
          overrides: [],
        },
      },
      {} as never,
    );
    expect(input).not.toContain("<Override");
    expect(input).toContain('Extension="xml" ContentType="application/xml"');
  });

  it("keeps a covering source binary default instead of adding a media override", () => {
    const input = finalizeContentTypes(
      { "word/media/image1.bin": new Uint8Array([1]) },
      {
        resolve: () => undefined,
        mediaContentTypes: {
          bin: "application/vnd.openxmlformats-officedocument.oleObject",
        },
        source: {
          preserveSourceDeclarations: true,
          defaults: [{ extension: "bin", contentType: "image/x-wmf" }],
          overrides: [],
        },
      },
      {} as never,
    );
    expect(input).not.toContain("<Override");
    expect(input).toContain('Extension="bin" ContentType="image/x-wmf"');
  });
});
