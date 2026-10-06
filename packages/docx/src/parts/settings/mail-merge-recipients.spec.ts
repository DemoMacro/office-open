import { unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateDocumentSync } from "../../generate";
import { parseDocumentSync } from "../../parse";

const decoder = new TextDecoder();

describe("mail merge recipients part", () => {
  it("packs recipients, wires settings, and round-trips", () => {
    const options = {
      sections: [{ children: [] }],
      mailMergeRecipients: [{ recipients: [{ active: true, column: 1, uniqueTag: "AA==" }] }],
    };
    const output = generateDocumentSync(options, { type: "uint8array" });
    const files = unzipSync(output);
    expect(decoder.decode(files["word/recipients1.xml"]!)).toContain("<w:recipientData>");
    expect(decoder.decode(files["word/_rels/settings.xml.rels"]!)).toContain(
      'Target="recipients1.xml"',
    );
    expect(decoder.decode(files["[Content_Types].xml"]!)).toContain("mailMergeRecipientData+xml");

    const parsed = parseDocumentSync(output);
    expect(parsed.mailMergeRecipients).toEqual(options.mailMergeRecipients);
    expect(parsed.settings?.mailMerge?.odso?.recipientData).toEqual(["rId1"]);
  });

  it("preserves an external mail merge data source relationship", () => {
    const source = zipSync({
      "word/document.xml": new TextEncoder().encode(
        '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
      ),
      "word/settings.xml": new TextEncoder().encode(
        '<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<w:mailMerge><w:mainDocumentType w:val="letters"/><w:dataType w:val="database"/>' +
          '<w:dataSource r:id="rId1"/></w:mailMerge></w:settings>',
      ),
      "word/_rels/settings.xml.rels": new TextEncoder().encode(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailData" ' +
          'Target="https://example.invalid/data.mdb" TargetMode="External"/></Relationships>',
      ),
    });

    const parsed = parseDocumentSync(source);
    expect(parsed.settings?.mailMerge?.dataSource).toBe("rId1");
    const output = unzipSync(generateDocumentSync(parsed, { type: "uint8array" }));
    const settings = new TextDecoder().decode(output["word/settings.xml"]!);
    const rels = new TextDecoder().decode(output["word/_rels/settings.xml.rels"]!);
    expect(settings).toContain('<w:dataSource r:id="rId1"/>');
    expect(rels).toContain(
      'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailData"',
    );
    expect(rels).toContain('Target="https://example.invalid/data.mdb" TargetMode="External"');
    expect(Object.keys(output)).not.toContain("word/recipients1.xml");
  });
});
