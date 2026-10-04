import { unzipSync } from "@office-open/core";
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
});
