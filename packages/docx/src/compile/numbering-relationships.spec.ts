import { strFromU8, unzipSync, zipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vitest";

import { generateDocumentSync } from "../generate";
import { parseDocumentSync } from "../parse";

describe("numbering relationships", () => {
  it("preserves unused source image relationships", () => {
    const xml = (value: string): Uint8Array => new TextEncoder().encode(value);
    const source = zipSync({
      "[Content_Types].xml": xml(
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Default Extension="png" ContentType="image/png"/>' +
          '<Default Extension="gif" ContentType="image/gif"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
          "</Types>",
      ),
      "_rels/.rels": xml(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
          "</Relationships>",
      ),
      "word/document.xml": xml(
        '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
      ),
      "word/_rels/document.xml.rels": xml(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
          "</Relationships>",
      ),
      "word/numbering.xml": xml(
        '<?xml version="1.0"?><w:numbering xmlns:o="urn:schemas-microsoft-com:office:office" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
          'xmlns:v="urn:schemas-microsoft-com:vml" ' +
          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:numPicBullet w:numPicBulletId="1"><w:pict><v:shape style="width:10pt;height:10pt">' +
          '<v:imagedata r:id="rId2" o:title="bullet"/></v:shape></w:pict></w:numPicBullet></w:numbering>',
      ),
      "word/_rels/numbering.xml.rels": xml(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.gif"/>' +
          "</Relationships>",
      ),
      "word/media/image1.png": new Uint8Array([1]),
      "word/media/image2.gif": new Uint8Array([2]),
    });

    const parsed = parseDocumentSync(source);
    const output = unzipSync(generateDocumentSync(parsed, { type: "uint8array" }));
    const rels = strFromU8(output["word/_rels/numbering.xml.rels"]!);
    expect(rels).toContain('Id="rId1"');
    expect(rels).toContain('Target="media/image1.png"');
    expect(rels).toContain('Id="rId2"');
    expect(rels).toContain('Target="media/image2.gif"');
    expect(strFromU8(output["word/numbering.xml"]!)).toContain('<v:imagedata r:id="rId2"');
  });
});
