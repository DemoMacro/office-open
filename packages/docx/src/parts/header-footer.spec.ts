import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

/** Unzip a generated package and decode one part as UTF-8 text. */
function decodePart(output: Uint8Array, path: string): string {
  const unzipped = unzipSync(output);
  const entry = unzipped[path];
  expect(entry, `${path} should be zipped`).toBeDefined();
  return new TextDecoder().decode(entry);
}

describe("header/footer reference id stability", () => {
  /** Map of `headerReference:default` style keys to their `r:id` values. */
  function referenceIds(partXml: string): Record<string, string> {
    const refs: Record<string, string> = {};
    for (const m of partXml.matchAll(/<w:(headerReference|footerReference)\s[^>]*\/>/g)) {
      const type = /w:type="([^"]+)"/.exec(m[0])?.[1];
      const rId = /r:id="([^"]+)"/.exec(m[0])?.[1];
      if (type && rId) refs[`${m[1]}:${type}`] = rId;
    }
    return refs;
  }

  const doc = {
    sections: [
      {
        children: [{ paragraph: { children: [{ text: "Body" }] } }],
        headers: { default: [{ paragraph: { children: [{ text: "Header" }] } }] },
        footers: { even: [{ paragraph: { children: [{ text: "Footer" }] } }] },
      },
    ],
  };

  it("keeps the reference ids of an opened document stable across saves", async () => {
    const first = await generateDocument(doc, { type: "uint8array" });
    const second = await generateDocument(parseDocumentSync(first), { type: "uint8array" });
    const third = await generateDocument(parseDocumentSync(second), { type: "uint8array" });

    const firstIds = referenceIds(decodePart(first, "word/document.xml"));
    const secondIds = referenceIds(decodePart(second, "word/document.xml"));
    const thirdIds = referenceIds(decodePart(third, "word/document.xml"));
    expect(secondIds).toEqual(firstIds);
    expect(thirdIds).toEqual(secondIds);
  });

  it("reuses one relationship when two slots reference the same part", async () => {
    // Source shape: the default and even slots both point at header1.xml. The
    // shared part needs one relationship and one id, not a second allocation
    // that would duplicate the id or churn on every save.
    const shared = {
      sections: [
        {
          children: [{ paragraph: { children: [{ text: "Body" }] } }],
          headers: {
            default: [{ paragraph: { children: [{ text: "Shared" }] } }],
            even: [{ paragraph: { children: [{ text: "Shared" }] } }],
            partNames: { default: "header1.xml", even: "header1.xml" },
          },
        },
      ],
    };

    const output = await generateDocument(shared, { type: "uint8array" });
    const refs = referenceIds(decodePart(output, "word/document.xml"));
    expect(refs["headerReference:default"]).toBeDefined();
    expect(refs["headerReference:even"]).toBe(refs["headerReference:default"]);

    const rels = decodePart(output, "word/_rels/document.xml.rels");
    expect(rels.match(/Target="header1\.xml"/g)).toHaveLength(1);
  });
});

describe("header/footer part naming", () => {
  const doc = {
    sections: [
      {
        children: [{ paragraph: { children: [{ text: "Body" }] } }],
        headers: {
          even: [{ paragraph: { children: [{ text: "Even header" }] } }],
          default: [{ paragraph: { children: [{ text: "Default header" }] } }],
        },
      },
    ],
  };

  it("parses back the source part name per slot (round-trip pin)", async () => {
    const output = await generateDocument(doc, { type: "uint8array" });
    const opts = parseDocumentSync(output);

    const headers = opts.sections?.[0]?.headers;
    expect(headers?.partNames).toEqual({
      default: "header1.xml",
      even: "header2.xml",
    });
  });

  it("keeps content in the same part file when round-tripped", async () => {
    const first = await generateDocument(doc, { type: "uint8array" });
    const second = await generateDocument(parseDocumentSync(first), { type: "uint8array" });

    // Content stays in its source part instead of sliding to the next file
    // when the slot iteration order differs from the numbering.
    expect(decodePart(second, "word/header1.xml")).toContain("Default header");
    expect(decodePart(second, "word/header2.xml")).toContain("Even header");
  });
});
