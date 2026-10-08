import { unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../../../generate";
import { parseDocumentSync } from "../../../parse";
import type { DocumentOptions } from "../../../parts";
import type { RunOptions } from "../../../parts/paragraph/run";

const ENCODER = new TextEncoder();
const DECODER = new TextDecoder();

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:synthetic="https://example.invalid/synthetic-drawing">
  <w:body>
    <w:p>
      <w:r w:rsidR="00AA0001">
        <w:t>before</w:t>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="First synthetic"/>
            <a:graphic><a:graphicData uri="https://example.invalid/synthetic-drawing">
              <synthetic:shape id="2" state="first"/>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
        <w:t>middle</w:t>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="3" name="Second synthetic"/>
            <a:graphic><a:graphicData uri="https://example.invalid/synthetic-drawing">
              <synthetic:shape id="4" state="second"/>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

function source() {
  return zipSync({
    "[Content_Types].xml": ENCODER.encode(CONTENT_TYPES_XML),
    "word/document.xml": ENCODER.encode(DOCUMENT_XML),
  });
}

function mixedRun(options: DocumentOptions): RunOptions {
  const child = options.sections?.[0]?.children?.[0];
  if (!child || !("paragraph" in child) || typeof child.paragraph === "string")
    throw new Error("fixture has no paragraph options");
  const run = child.paragraph.children?.[0];
  if (!run || typeof run === "string" || !("children" in run))
    throw new Error("mixed drawing run was split");
  return run;
}

describe("mixed DrawingML runs", () => {
  it("keeps text and multiple drawings inside one source run", async () => {
    const parsed = parseDocumentSync(source()) as DocumentOptions;
    const run = mixedRun(parsed);
    expect(run.additionRsid).toBe("00AA0001");
    expect(run.children).toHaveLength(4);
    expect(run.children?.[0]).toBe("before");
    expect(run.children?.[1]).toMatchObject({
      unsupportedDrawing: { graphicDataUri: "https://example.invalid/synthetic-drawing" },
    });
    expect(run.children?.[2]).toBe("middle");
    expect(run.children?.[3]).toMatchObject({
      unsupportedDrawing: {
        graphicData: {
          elements: [
            {
              name: "synthetic:shape",
              attributes: { id: 4, state: "second" },
            },
          ],
        },
      },
    });

    const first = await generateDocument(parsed, { type: "uint8array" });
    const xml = DECODER.decode(unzipSync(first)["word/document.xml"]!);
    expect(xml).toContain("<w:t>before</w:t>");
    expect(xml).toContain('name="First synthetic"');
    expect(xml).toContain("<w:t>middle</w:t>");
    expect(xml).toContain('name="Second synthetic"');

    const reparsed = parseDocumentSync(first) as DocumentOptions;
    expect(await generateDocument(reparsed, { type: "uint8array" })).toEqual(first);
  });
});
