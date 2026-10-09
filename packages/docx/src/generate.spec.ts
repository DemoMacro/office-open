import { unzipSync, zipSync } from "@office-open/core";
import type { GroupChildMediaData } from "@shared/media";
import { describe, expect, it } from "vite-plus/test";

import { compileDocument } from "./compiler";
import { generateDocument, generateDocumentSync } from "./generate";
import { parseDocumentSync } from "./parse";

describe("generateDocument entry guards", () => {
  it("names the missing sections array instead of dying in the compiler", () => {
    expect(() => generateDocument({} as never)).toThrow(/sections is required/);
  });
});

describe("raw cross-paragraph section hosts", () => {
  it("does not append another section break after hosted rawXml", () => {
    const xml = new TextDecoder().decode(
      compileDocument({
        sections: [
          {
            children: [
              {
                rawXml:
                  '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>' +
                  '<w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>',
              },
            ],
          },
          { children: [{ paragraph: { children: ["Next"] } }] },
        ],
      })["word/document.xml"] as Uint8Array,
    );
    expect(xml.match(/<w:sectPr[ />]/g)).toHaveLength(2);
    expect(xml).not.toContain("><w:p><w:pPr><w:sectPr/>");
  });
});

describe("package variants", () => {
  const variants = [
    {
      variant: "standard",
      mainContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
    {
      variant: "macro",
      mainContentType: "application/vnd.ms-word.document.macroEnabled.main+xml",
      mimeType: "application/vnd.ms-word.document.macroEnabled.12",
    },
    {
      variant: "template",
      mainContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    },
    {
      variant: "macroTemplate",
      mainContentType: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
      mimeType: "application/vnd.ms-word.template.macroEnabled.12",
    },
  ] as const;
  const encrypted = {
    sections: [],
    encrypted: { data: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]) },
  };

  it("forces the main document content type for each output variant", () => {
    for (const { variant, mainContentType } of variants) {
      const files = compileDocument(
        {
          sections: [{ children: [{ paragraph: { children: ["Variant"] } }] }],
        },
        [],
        0,
        undefined,
        variant,
      );
      const xml = new TextDecoder().decode(files["[Content_Types].xml"] as Uint8Array);
      expect(xml).toContain(`PartName="/word/document.xml" ContentType="${mainContentType}"`);
    }
  });

  it("uses the variant MIME for generated and encrypted passthrough blobs", async () => {
    for (const { variant, mimeType } of variants) {
      const generated = await generateDocument(
        { sections: [{ children: [{ paragraph: { children: ["Variant"] } }] }] },
        { type: "blob", packageVariant: variant },
      );
      const passthrough = generateDocumentSync(encrypted, {
        type: "blob",
        packageVariant: variant,
      });
      expect(generated.type).toBe(mimeType.toLowerCase());
      expect(passthrough.type).toBe(mimeType.toLowerCase());
    }
  });
});

describe("chart embedding rels", () => {
  it("emits the c:externalData rel with a quoted package relationship type", () => {
    const files = compileDocument({
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    chart: {
                      type: "column",
                      series: [{ values: [1, 2, 3] }],
                      transformation: { width: 5486400, height: 3200400 },
                      externalData: {
                        relationshipId: "rId1",
                        fileName: "Chart.xlsx",
                        data: new Uint8Array([1, 2, 3]),
                      },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const rels = files["word/charts/_rels/chart1.xml.rels"];
    expect(rels).toBeDefined();
    const xml = new TextDecoder().decode(rels as Uint8Array);
    expect(xml).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"',
    );
    expect(xml).toContain('Target="../embeddings/Chart.xlsx"');
  });

  it("claims source chart style and color relationships", () => {
    const files = compileDocument({
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    chart: {
                      type: "column",
                      series: [{ values: [1, 2, 3] }],
                      transformation: { width: 5486400, height: 3200400 },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
      rawParts: [{ path: "word/charts/style1.xml", data: "<chartStyle/>" }],
      passthroughRelationships: [
        {
          source: "word/charts/chart1.xml",
          relationshipType: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
          target: "style1.xml",
          rId: "rId1",
        },
      ],
    });
    const rels = new TextDecoder().decode(files["word/charts/_rels/chart1.xml.rels"] as Uint8Array);
    expect(rels).toContain(
      'Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle"',
    );
    expect(rels).toContain('Target="style1.xml"');
  });
});

describe("package metadata presence", () => {
  const contentTypes = {
    defaults: [],
    overrides: [
      {
        partName: "/word/document.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      },
    ],
  };

  it("omits source-absent core and app metadata", () => {
    const files = compileDocument({ sections: [], contentTypes });
    expect(Object.keys(files)).not.toContain("docProps/core.xml");
    expect(Object.keys(files)).not.toContain("docProps/app.xml");
    expect(Object.keys(files)).not.toContain("word/styles.xml");
    expect(Object.keys(files)).not.toContain("word/settings.xml");
    const rels = new TextDecoder().decode(files["_rels/.rels"] as Uint8Array);
    expect(rels).not.toContain("docProps/core.xml");
    expect(rels).not.toContain("docProps/app.xml");
    const documentRels = new TextDecoder().decode(
      files["word/_rels/document.xml.rels"] as Uint8Array,
    );
    expect(documentRels).not.toContain("styles.xml");
    expect(documentRels).not.toContain("settings.xml");
  });

  it("keeps source-backed metadata and explicit app metadata", () => {
    const files = compileDocument({
      sections: [],
      contentTypes: {
        ...contentTypes,
        overrides: [
          ...contentTypes.overrides,
          {
            partName: "/docProps/core.xml",
            contentType: "application/vnd.openxmlformats-package.core-properties+xml",
          },
        ],
      },
      appProperties: { application: "Test" },
    });
    expect(Object.keys(files)).toContain("docProps/core.xml");
    expect(Object.keys(files)).toContain("docProps/app.xml");
    const rels = new TextDecoder().decode(files["_rels/.rels"] as Uint8Array);
    expect(rels).toContain("docProps/core.xml");
    expect(rels).toContain("docProps/app.xml");
  });

  it("preserves orphan metadata declarations from damaged sources", () => {
    const source = zipSync({
      "[Content_Types].xml": new TextEncoder().encode(
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
          "</Types>",
      ),
      "_rels/.rels": new TextEncoder().encode(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
          '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
          "</Relationships>",
      ),
      "word/document.xml": new TextEncoder().encode(
        '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
      ),
    });
    const parsed = parseDocumentSync(source);
    expect(
      parsed.contentTypes?.overrides.some((override) => override.partName === "/docProps/core.xml"),
    ).toBe(true);
    expect(
      parsed.contentTypes?.overrides.some((override) => override.partName === "/docProps/app.xml"),
    ).toBe(true);
    const output = generateDocumentSync(parsed, { type: "uint8array" });
    const rootRelsXml = new TextDecoder().decode(unzipSync(output)["_rels/.rels"]!);
    expect(rootRelsXml).toContain("/relationships/officeDocument");
    expect(rootRelsXml).toContain("docProps/core.xml");
    expect(rootRelsXml).toContain("docProps/app.xml");
    const contentTypesXml = new TextDecoder().decode(unzipSync(output)["[Content_Types].xml"]!);
    expect(contentTypesXml).toContain("/docProps/core.xml");
    expect(contentTypesXml).toContain("/docProps/app.xml");
  });

  it("round-trips a non-default primary document path", () => {
    const source = zipSync({
      "[Content_Types].xml": new TextEncoder().encode(
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Override PartName="/word/document2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          "</Types>",
      ),
      "_rels/.rels": new TextEncoder().encode(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document2.xml"/>' +
          "</Relationships>",
      ),
      "word/document2.xml": new TextEncoder().encode(
        '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
      ),
    });
    const parsed = parseDocumentSync(source);
    expect(parsed.primaryPartPath).toBe("word/document2.xml");
    const output = unzipSync(generateDocumentSync(parsed, { type: "uint8array" }));
    expect(Object.keys(output)).toContain("word/document2.xml");
    expect(Object.keys(output)).toContain("word/_rels/document2.xml.rels");
    expect(new TextDecoder().decode(output["_rels/.rels"]!)).toContain(
      'Target="word/document2.xml"',
    );
    expect(new TextDecoder().decode(output["[Content_Types].xml"]!)).toContain(
      'PartName="/word/document2.xml"',
    );
  });
});

describe("picture media dedup", () => {
  it("keeps per-reference extent for byte-identical images", () => {
    const bytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const emu = (px: number) => px * 9525;
    const files = compileDocument({
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    picture: {
                      type: "png",
                      data: bytes,
                      transformation: { width: emu(166), height: emu(150) },
                    },
                  },
                  {
                    picture: {
                      type: "png",
                      data: bytes,
                      transformation: { width: emu(140), height: emu(93) },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const xml = new TextDecoder().decode(files["word/document.xml"] as Uint8Array);
    const extents = [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map(
      (m) => `${m[1]}x${m[2]}`,
    );
    expect(extents).toEqual([`${emu(166)}x${emu(150)}`, `${emu(140)}x${emu(93)}`]);
  });

  it("keeps per-reference svg fallback for byte-identical svg data", () => {
    const svgBytes = new Uint8Array([60, 115, 118, 103, 62]);
    const emu = (px: number) => px * 9525;
    const files = compileDocument({
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    picture: {
                      type: "svg",
                      data: svgBytes,
                      fallback: {
                        type: "png",
                        data: new Uint8Array([10, 11, 12]),
                        fileName: "small.png",
                      },
                      transformation: { width: emu(50), height: emu(50) },
                    },
                  },
                  {
                    picture: {
                      type: "svg",
                      data: svgBytes,
                      fallback: {
                        type: "png",
                        data: new Uint8Array([20, 21, 22, 23]),
                        fileName: "large.png",
                      },
                      transformation: { width: emu(80), height: emu(80) },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const xml = new TextDecoder().decode(files["word/document.xml"] as Uint8Array);
    const relsXml = new TextDecoder().decode(files["word/_rels/document.xml.rels"] as Uint8Array);
    const targetOf = new Map(
      [...relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1]!, m[2]!]),
    );
    const fallbackTargets = [...xml.matchAll(/<a:blip r:embed="([^"]+)"/g)].map(
      (m) => targetOf.get(m[1]!) ?? "",
    );
    // Each svg reference resolves its OWN raster fallback — the second
    // reference must not inherit the first registrant's entry.
    expect(fallbackTargets.some((t) => t.endsWith("small.png"))).toBe(true);
    expect(fallbackTargets.some((t) => t.endsWith("large.png"))).toBe(true);
  });

  it("resolves the blip rel of a fresh picture child inside a wpg group", () => {
    // A fresh-authored group picture carries no fileName — the media
    // registration must adopt the allocated name so the {fileName} placeholder
    // resolves (an unresolved placeholder emits r:embed="{undefined}").
    const emu = (px: number) => px * 9525;
    const files = compileDocument({
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    wpgGroup: {
                      transformation: { width: emu(254), height: emu(128) },
                      childOffsetX: 0,
                      childOffsetY: 0,
                      childExtentWidth: emu(254),
                      childExtentHeight: emu(128),
                      children: [
                        // A fresh-authored child carries no fileName (the type
                        // demands one, but fresh authoring can't know it) —
                        // registration must allocate.
                        {
                          type: "png",
                          data: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2]),
                          transformation: {
                            offset: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
                            pixels: { x: 100, y: 100 },
                            emus: { x: emu(100), y: emu(100) },
                          },
                        } as GroupChildMediaData,
                      ],
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    });
    const xml = new TextDecoder().decode(files["word/document.xml"] as Uint8Array);
    const relsXml = new TextDecoder().decode(files["word/_rels/document.xml.rels"] as Uint8Array);
    const embeds = [...xml.matchAll(/<a:blip r:embed="([^"]+)"/g)].map((m) => m[1]!);
    expect(embeds).toHaveLength(1);
    const targetOf = new Map(
      [...relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1]!, m[2]!]),
    );
    // The embed id resolves to a media part, not the literal placeholder.
    expect(targetOf.get(embeds[0]!)).toMatch(/^media\/image\d+\.png$/);
  });

  it("keeps typed VML background media pristine across generate runs", () => {
    const bytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const options = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    picture: {
                      type: "png",
                      data: bytes,
                      transformation: { width: 9525, height: 9525 },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
      background: {
        vmlBackground: {
          fill: { type: "frame", relationshipId: "{bg.png}" },
        },
        rawMedia: [{ fileName: "bg.png", type: "png", data: bytes }],
      },
    } as Parameters<typeof compileDocument>[0];
    compileDocument(options);
    // The first run dedups bg.png against the body picture — the rename must
    // stay inside that run, never land on the caller's options object.
    expect(options.background?.vmlBackground?.fill?.relationshipId).toBe("{bg.png}");
    const second = compileDocument(options);
    const xml = new TextDecoder().decode(second["word/document.xml"] as Uint8Array);
    expect(xml).not.toMatch(/\{[a-zA-Z0-9_.]+\}/);

    const generated = generateDocumentSync(options);
    const generatedXml = new TextDecoder().decode(
      unzipSync(generated)["word/document.xml"] as Uint8Array,
    );
    expect(generatedXml).not.toMatch(/\{[a-zA-Z0-9_.]+\}/);
    const parsed = parseDocumentSync(generated);
    expect(parsed.background?.vmlBackground?.fill?.relationshipId).toBe("{bg.png}");
    expect(parsed.background?.rawMedia).toHaveLength(1);
  });

  it("re-registers group chart children on a reused options object", () => {
    const mediaTransformation = {
      emus: { x: 952500, y: 952500 },
      pixels: { x: 100, y: 100 },
      offset: { emus: { x: 0, y: 0 }, pixels: { x: 0, y: 0 } },
    };
    const options = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    wpgGroup: {
                      children: [
                        {
                          type: "chart",
                          transformation: mediaTransformation,
                          chartOptions: { type: "column", series: [{ values: [1, 2, 3] }] },
                        },
                      ],
                      transformation: { width: 1905000, height: 1905000 },
                      childExtentWidth: 952500,
                      childExtentHeight: 952500,
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    } as Parameters<typeof compileDocument>[0];
    compileDocument(options);
    const second = compileDocument(options);
    const xml = new TextDecoder().decode(second["word/document.xml"] as Uint8Array);
    // A persisted chart key from the first run must not skip the second run's
    // chart registration — the placeholder has to resolve to a real part.
    expect(xml).not.toMatch(/\{chart:/);
    expect(Object.keys(second).some((f) => f.startsWith("word/charts/"))).toBe(true);
  });
});

describe("chart and smartart placeholders in header/footer/notes parts", () => {
  it("resolves {chart:key} in a header against the header's own rels", () => {
    const files = compileDocument({
      sections: [
        {
          headers: {
            default: [
              {
                paragraph: {
                  children: [
                    {
                      chart: {
                        type: "column",
                        series: [{ values: [1, 2, 3] }],
                        transformation: { width: 5486400, height: 3200400 },
                      },
                    },
                  ],
                },
              },
            ],
          },
          children: [{ paragraph: { children: ["Filler line for the body."] } }],
        },
      ],
    });
    const headerXml = new TextDecoder().decode(files["word/header1.xml"] as Uint8Array);
    expect(headerXml).not.toMatch(/\{chart:/);
    const headerRels = new TextDecoder().decode(files["word/_rels/header1.xml.rels"] as Uint8Array);
    expect(headerRels).toContain('Target="charts/chart1.xml"');
    expect(files["word/charts/chart1.xml"]).toBeDefined();
  });

  it("resolves {smartart:*:key} in a footer against the footer's own rels", () => {
    const files = compileDocument({
      sections: [
        {
          footers: {
            default: [
              {
                paragraph: {
                  children: [
                    {
                      smartArt: {
                        nodes: [{ text: "Alpha" }, { text: "Beta" }],
                        transformation: { width: 5486400, height: 3200400 },
                      },
                    },
                  ],
                },
              },
            ],
          },
          children: [{ paragraph: { children: ["Filler line for the body."] } }],
        },
      ],
    });
    const footerXml = new TextDecoder().decode(files["word/footer1.xml"] as Uint8Array);
    expect(footerXml).not.toMatch(/\{smartart/);
    const footerRels = new TextDecoder().decode(files["word/_rels/footer1.xml.rels"] as Uint8Array);
    for (const diagram of ["data1", "layout1", "quickStyle1", "colors1"]) {
      expect(footerRels).toContain(`Target="diagrams/${diagram}.xml"`);
    }
    expect(files["word/diagrams/data1.xml"]).toBeDefined();
  });
});

describe("sparse drawing part paths", () => {
  function movePart(
    input: Uint8Array,
    from: string,
    to: string,
    ownerPath: string,
    previousTarget: string,
    newTarget: string,
  ): Uint8Array {
    const files = unzipSync(input);
    const part = files[from];
    expect(part, from).toBeDefined();
    delete files[from];
    files[to] = part!;
    const relsPath = `word/_rels/${ownerPath.slice("word/".length)}.rels`;
    const rels = new TextDecoder().decode(files[relsPath]!);
    files[relsPath] = new TextEncoder().encode(
      rels.replace(`Target="${previousTarget}"`, `Target="${newTarget}"`),
    );
    const contentTypes = new TextDecoder().decode(files["[Content_Types].xml"]!);
    files["[Content_Types].xml"] = new TextEncoder().encode(
      contentTypes.replace(`PartName="/${from}"`, `PartName="/${to}"`),
    );
    return zipSync(files);
  }

  function moveDocumentPart(
    input: Uint8Array,
    from: string,
    to: string,
    previousTarget: string,
  ): Uint8Array {
    return movePart(input, from, to, "word/document.xml", previousTarget, to.slice("word/".length));
  }

  it("keeps a consumed chart at its source package path", async () => {
    const fresh = await generateDocument(
      {
        sections: [
          {
            children: [
              {
                paragraph: {
                  children: [
                    {
                      chart: {
                        type: "column",
                        series: [{ values: [1, 2, 3] }],
                        transformation: { width: 5486400, height: 3200400 },
                      },
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    );
    const source = moveDocumentPart(
      fresh,
      "word/charts/chart1.xml",
      "word/custom/chart.xml",
      "charts/chart1.xml",
    );
    expect(new TextDecoder().decode(unzipSync(source)["word/_rels/document.xml.rels"]!)).toContain(
      'Target="custom/chart.xml"',
    );
    const parsed = parseDocumentSync(source);
    const regenerated = await generateDocument(parsed, { type: "uint8array" });
    const files = unzipSync(regenerated);
    expect(files["word/custom/chart.xml"]).toBeDefined();
    expect(files["word/charts/chart1.xml"]).toBeUndefined();
    expect(
      (parsed.rawParts ?? []).some((part) => part.path === "word/custom/chart.xml"),
      JSON.stringify(parsed.rawParts),
    ).toBe(false);
    expect(new TextDecoder().decode(files["word/_rels/document.xml.rels"]!)).toContain(
      'Target="custom/chart.xml"',
    );
  });

  it("keeps consumed SmartArt data at its source package path", async () => {
    const fresh = await generateDocument(
      {
        sections: [
          {
            children: [
              {
                paragraph: {
                  children: [
                    {
                      smartArt: {
                        nodes: [{ text: "Alpha" }],
                        transformation: { width: 5486400, height: 3200400 },
                      },
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    );
    const source = moveDocumentPart(
      fresh,
      "word/diagrams/data1.xml",
      "word/custom/data.xml",
      "diagrams/data1.xml",
    );
    const parsed = parseDocumentSync(source);
    const regenerated = await generateDocument(parsed, { type: "uint8array" });
    const files = unzipSync(regenerated);
    expect(files["word/custom/data.xml"]).toBeDefined();
    expect(files["word/diagrams/data1.xml"]).toBeUndefined();
    expect(
      (parsed.rawParts ?? []).some((part) => part.path === "word/custom/data.xml"),
      JSON.stringify(parsed.rawParts),
    ).toBe(false);
    expect(new TextDecoder().decode(files["word/_rels/document.xml.rels"]!)).toContain(
      'Target="custom/data.xml"',
    );
  });
});
