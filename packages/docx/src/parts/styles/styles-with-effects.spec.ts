import { strFromU8, unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { generateDocumentSync } from "../../generate";
import { parseDocumentSync } from "../../parse";

describe("stylesWithEffects part", () => {
  it("round-trips the legacy companion without rawParts", () => {
    const xml = (value: string) => new TextEncoder().encode(value);
    const source = zipSync({
      "[Content_Types].xml": xml(
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>' +
          "</Types>",
      ),
      "_rels/.rels": xml(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      ),
      "word/document.xml": xml(
        '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
      ),
      "word/_rels/document.xml.rels": xml(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/>' +
          "</Relationships>",
      ),
      "word/stylesWithEffects.xml": xml(
        '<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:style w:type="paragraph" w:styleId="Effects"><w:name w:val="Effects"/></w:style></w:styles>',
      ),
    });

    const parsed = parseDocumentSync(source);
    expect(parsed.rawParts).toBeUndefined();
    expect(parsed.stylesWithEffects?.paragraphStyles).toHaveLength(1);

    const output = unzipSync(generateDocumentSync(parsed, { type: "uint8array" }));
    expect(output["word/stylesWithEffects.xml"]).toBeDefined();
    expect(strFromU8(output["word/_rels/document.xml.rels"]!)).toContain("stylesWithEffects");
    expect(strFromU8(output["[Content_Types].xml"]!)).toContain("/word/stylesWithEffects.xml");
  });
});
