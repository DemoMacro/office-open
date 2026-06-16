import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ContentTypesInput } from "./contenttypes";
import { buildContentTypesFromRegistry, contentTypesDesc } from "./contenttypes";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: ContentTypesInput) {
  const xml = contentTypesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return contentTypesDesc.parse(el, readCtx);
}

describe("contentTypesDesc round-trip", () => {
  it("round-trips defaults", () => {
    const result = roundTrip({
      defaults: [
        { extension: "png", contentType: "image/png" },
        { extension: "xml", contentType: "application/xml" },
      ],
      overrides: [],
    });
    expect(result.defaults).toHaveLength(2);
    expect(result.defaults[0].extension).toBe("png");
    expect(result.defaults[0].contentType).toBe("image/png");
    expect(result.defaults[1].extension).toBe("xml");
  });

  it("round-trips overrides", () => {
    const result = roundTrip({
      defaults: [],
      overrides: [
        {
          partName: "/word/document.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        },
      ],
    });
    expect(result.overrides).toHaveLength(1);
    expect(result.overrides[0].partName).toBe("/word/document.xml");
    expect(result.overrides[0].contentType).toContain("wordprocessingml");
  });

  it("round-trips both defaults and overrides", () => {
    const result = roundTrip({
      defaults: [
        {
          extension: "rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
      ],
      overrides: [
        {
          partName: "/word/styles.xml",
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        },
        {
          partName: "/word/settings.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        },
      ],
    });
    expect(result.defaults).toHaveLength(1);
    expect(result.overrides).toHaveLength(2);
  });

  it("round-trips empty content types", () => {
    const result = roundTrip({ defaults: [], overrides: [] });
    expect(result.defaults).toHaveLength(0);
    expect(result.overrides).toHaveLength(0);
  });
});

describe("buildContentTypesFromRegistry", () => {
  const partNames = (r: ContentTypesInput) => new Set(r.overrides.map((o) => o.partName));

  it("emits fresh-compile static parts when freshCompile is set", () => {
    const names = partNames(buildContentTypesFromRegistry(new Map([["freshCompile", true]])));
    // document is `always`; the rest are fresh-compile conditionals.
    expect(names).toContain("/word/document.xml");
    expect(names).toContain("/word/styles.xml");
    expect(names).toContain("/word/numbering.xml");
    expect(names).toContain("/word/footnotes.xml");
    expect(names).toContain("/word/endnotes.xml");
    expect(names).toContain("/word/settings.xml");
    expect(names).toContain("/word/fontTable.xml");
    expect(names).toContain("/docProps/core.xml");
    expect(names).toContain("/docProps/app.xml");
    expect(names).toContain("/docProps/custom.xml");
    // comments absent until hasComments
    expect(names).not.toContain("/word/comments.xml");
  });

  it("keeps only the always-part when freshCompile is unset", () => {
    const names = partNames(buildContentTypesFromRegistry(new Map()));
    expect(names).toContain("/word/document.xml");
    expect(names).not.toContain("/word/styles.xml");
    expect(names).not.toContain("/docProps/core.xml");
  });

  it("toggles the comments Override with hasComments", () => {
    const on = partNames(
      buildContentTypesFromRegistry(
        new Map([
          ["freshCompile", true],
          ["hasComments", true],
        ]),
      ),
    );
    const off = partNames(buildContentTypesFromRegistry(new Map([["freshCompile", true]])));
    expect(on).toContain("/word/comments.xml");
    expect(off).not.toContain("/word/comments.xml");
  });

  it("expands repeated parts by their count fact", () => {
    const names = partNames(
      buildContentTypesFromRegistry(
        new Map<string, boolean | number>([
          ["freshCompile", true],
          ["headerCount", 2],
          ["footerCount", 1],
          ["chartCount", 1],
          ["smartArtCount", 1],
        ]),
      ),
    );
    expect(names).toContain("/word/header1.xml");
    expect(names).toContain("/word/header2.xml");
    expect(names).toContain("/word/footer1.xml");
    expect(names).toContain("/word/charts/chart1.xml");
    // SmartArt = five diagram parts per index.
    expect(names).toContain("/word/diagrams/data1.xml");
    expect(names).toContain("/word/diagrams/layout1.xml");
    expect(names).toContain("/word/diagrams/drawing1.xml");
  });

  it("appends dynamic altChunks and sub-documents", () => {
    const result = buildContentTypesFromRegistry(new Map([["freshCompile", true]]), {
      altChunks: [{ path: "/word/afchunks/chunk1.html", contentType: "text/html" }],
      subDocs: [{ path: "/word/subdoc1.xml" }],
    });
    const names = partNames(result);
    expect(names).toContain("/word/afchunks/chunk1.html");
    expect(names).toContain("/word/subdoc1.xml");
    const subDoc = result.overrides.find((o) => o.partName === "/word/subdoc1.xml");
    expect(subDoc?.contentType).toContain("wordprocessingml.document.main+xml");
  });

  it("always carries the standard media extension defaults", () => {
    const result = buildContentTypesFromRegistry(new Map([["freshCompile", true]]));
    const exts = new Set(result.defaults.map((d) => d.extension));
    expect(exts).toContain("png");
    expect(exts).toContain("xml");
    expect(exts).toContain("rels");
  });
});
