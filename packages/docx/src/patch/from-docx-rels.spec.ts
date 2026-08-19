import { strFromU8, unzipSync, zipSync } from "fflate";
// Real-archive regression specs for relationship id allocation — the fflate
// mocks in from-docx.spec.ts bypass the zip round-trip, so id collisions
// between patched-in hyperlinks/images and pre-existing relationship ids
// were never observable there (duplicate Id corrupts the package for Word).
import { describe, expect, it } from "vite-plus/test";

import { patchDocument } from "./from-docx";

const encode = (s: string) => new TextEncoder().encode(s);

const DOC_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<w:body><w:p><w:r><w:t>{{link_test}}</w:t></w:r></w:p>` +
  `<w:sectPr/></w:body></w:document>`;

const RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
  `</Relationships>`;

const CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `</Types>`;

const templateWithOccupiedRels = () =>
  zipSync({
    "word/document.xml": encode(DOC_XML),
    "word/_rels/document.xml.rels": encode(RELS_XML),
    "[Content_Types].xml": encode(CONTENT_TYPES_XML),
  });

describe("patchDocument relationship id allocation", () => {
  it("allocates hyperlink ids after the ids already in the rels", async () => {
    const output = await patchDocument({
      data: templateWithOccupiedRels(),
      outputType: "uint8array",
      placeholders: {
        link_test: {
          children: [{ hyperlink: { url: "https://example.com", children: ["Example"] } }],
          type: "paragraph",
        },
      },
    });

    const z = unzipSync(output);
    const rels = strFromU8(z["word/_rels/document.xml.rels"]!);
    const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]!);
    expect(new Set(ids).size).to.equal(ids.length);
    expect(ids).to.contain("rId3");
    expect(strFromU8(z["word/document.xml"]!)).to.contain('r:id="rId3"');
  });

  it("allocates image ids after a patched-in hyperlink on the same part", async () => {
    const png = Uint8Array.from([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0, 0,
      0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0a, 0x49, 0x44, 0x41,
      0x54, 0x78, 0x9c, 0x63, 0, 1, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45,
      0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]);
    const output = await patchDocument({
      data: templateWithOccupiedRels(),
      outputType: "uint8array",
      placeholders: {
        link_test: {
          children: [
            { hyperlink: { url: "https://example.com", children: ["Example"] } },
            {
              picture: {
                data: png,
                transformation: { height: 100, width: 100 },
                type: "png",
              },
            },
          ],
          type: "paragraph",
        },
      },
    });

    const z = unzipSync(output);
    const rels = strFromU8(z["word/_rels/document.xml.rels"]!);
    const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]!);
    expect(new Set(ids).size).to.equal(ids.length);
    expect(ids).to.contain("rId3"); // image (media flushes first)
    expect(ids).to.contain("rId4"); // hyperlink
    const document = strFromU8(z["word/document.xml"]!);
    expect(document).to.contain('r:embed="rId3"');
    expect(document).to.contain('r:id="rId4"');
  });
});
