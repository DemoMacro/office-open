import { unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { DocxParseError } from "../../../../errors";
import { generateDocumentSync } from "../../../../generate";
import { parseDocumentSync } from "../../../../parse";

const PRINTER_SETTINGS_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const PRINTER_DATA = new TextEncoder().encode("printer settings");

const documentXml = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<w:body>` +
  `<w:p><w:pPr><w:sectPr><w:printerSettings r:id="rId7"/></w:sectPr></w:pPr></w:p>` +
  `<w:sectPr><w:printerSettings r:id="rId7"/></w:sectPr>` +
  `</w:body></w:document>`;

const relsXml = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId7" Type="${PRINTER_SETTINGS_RELATIONSHIP_TYPE}" Target="printerSettings/printerSettings1.bin"/>` +
  `</Relationships>`;

const printerSettingsPackage = (includeBinary: boolean) => {
  const files: Record<string, Uint8Array> = {
    "word/document.xml": new TextEncoder().encode(documentXml()),
    "word/_rels/document.xml.rels": new TextEncoder().encode(relsXml()),
  };
  if (includeBinary) files["word/printerSettings/printerSettings1.bin"] = PRINTER_DATA;
  return zipSync(files);
};

const document = (zip: Record<string, Uint8Array>) =>
  new TextDecoder().decode(zip["word/document.xml"]!);

const relationships = (zip: Record<string, Uint8Array>) =>
  new TextDecoder().decode(zip["word/_rels/document.xml.rels"]!);

describe("section printerSettings round-trip", () => {
  it("preserves the resolved binary part path on parse", () => {
    const parsed = parseDocumentSync(printerSettingsPackage(true));
    expect(parsed.sections).toHaveLength(2);
    expect(parsed.sections[0]?.properties?.printerSettingsPath).toBe(
      "word/printerSettings/printerSettings1.bin",
    );
    expect(parsed.sections[1]?.properties?.printerSettingsPath).toBe(
      "word/printerSettings/printerSettings1.bin",
    );
  });

  it("registers one relationship for sections sharing printer settings", () => {
    const zip = unzipSync(
      generateDocumentSync(
        {
          sections: [
            {
              children: [],
              properties: { printerSettingsPath: "word/printerSettings/printerSettings1.bin" },
            },
            {
              children: [],
              properties: { printerSettingsPath: "word/printerSettings/printerSettings1.bin" },
            },
          ],
          rawParts: [
            {
              path: "word/printerSettings/printerSettings1.bin",
              data: PRINTER_DATA,
            },
          ],
        },
        { type: "uint8array" },
      ),
    );
    const rels = relationships(zip);
    expect(document(zip)).toContain("<w:printerSettings r:id=");
    expect(rels.match(new RegExp(PRINTER_SETTINGS_RELATIONSHIP_TYPE, "g"))).toHaveLength(1);
    expect(zip["word/printerSettings/printerSettings1.bin"]).toEqual(PRINTER_DATA);
  });

  it("throws a structured error when the referenced binary part is missing", () => {
    try {
      parseDocumentSync(printerSettingsPackage(false));
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(DocxParseError);
      const parseError = error as DocxParseError;
      expect(parseError.part).toBe("word/document.xml");
      expect(parseError.path).toBe("w:body/w:sectPr/w:printerSettings");
      expect(parseError.reason).toBe("missing-printer-settings-part");
      expect(parseError.targetPath).toBe("word/printerSettings/printerSettings1.bin");
    }
  });

  it("round-trips document references and the shared relationship", () => {
    const zip = unzipSync(
      generateDocumentSync(parseDocumentSync(printerSettingsPackage(true)), { type: "uint8array" }),
    );
    expect(document(zip).match(/<w:printerSettings r:id="/g)).toHaveLength(2);
    expect(
      relationships(zip).match(new RegExp(PRINTER_SETTINGS_RELATIONSHIP_TYPE, "g")),
    ).toHaveLength(1);
    expect(relationships(zip)).toContain('Target="printerSettings/printerSettings1.bin"');
    expect(zip["word/printerSettings/printerSettings1.bin"]).toEqual(PRINTER_DATA);
  });
});
