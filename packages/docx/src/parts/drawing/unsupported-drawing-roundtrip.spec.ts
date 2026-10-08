import { unzipSync, zipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../../generate";
import { parseDocumentSync } from "../../parse";
import type { DocumentOptions, ParagraphChild } from "../../parts";

const ENCODER = new TextEncoder();
const DECODER = new TextDecoder();

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:synthetic="https://example.invalid/synthetic-drawing">
  <w:body>
    <w:p>
      <w:r w:rsidR="00AA0001">
        <w:rPr><w:b/></w:rPr>
        <w:lastRenderedPageBreak/>
        <w:drawing xmlns:synthetic="https://example.invalid/synthetic-drawing">
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="7" name="Synthetic unsupported"/>
            <a:graphic><a:graphicData uri="https://example.invalid/synthetic-drawing">
              <synthetic:shape id="42" state="kept">
                <a:cNvPr id="8" name="Hover shape">
                  <a:hlinkHover r:id="rId2" tooltip="Synthetic hover"/>
                </a:cNvPr>
                <synthetic:data value="canonical"/>
              </synthetic:shape>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.invalid/hover" TargetMode="External"/>
</Relationships>`;

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const ALTERNATE_CONTENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/chartex"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:noProof/></w:rPr>
        <mc:AlternateContent>
          <mc:Choice Requires="cx1">
            <w:drawing>
              <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="5486400" cy="3200400"/>
                <wp:docPr id="1" name="Synthetic chart"/>
                <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2015/9/chartex">
                  <cx1:chart r:id="rId4"/>
                </a:graphicData></a:graphic>
              </wp:inline>
            </w:drawing>
          </mc:Choice>
          <mc:Fallback>
            <w:drawing>
              <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="5486400" cy="3200400"/>
                <wp:docPr id="1" name="Synthetic chart"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id="1" name="Synthetic chart"/><pic:cNvPicPr/></pic:nvPicPr>
                    <pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="3200400"/></a:xfrm>
                      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                  </pic:pic>
                </a:graphicData></a:graphic>
              </wp:inline>
            </w:drawing>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const ALTERNATE_CONTENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4"
    Type="http://schemas.microsoft.com/office/2014/relationships/chartEx"
    Target="charts/chartEx1.xml"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
</Relationships>`;

const GROUP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:synthetic="https://example.invalid/synthetic-shape">
  <w:body><w:p><w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="914400" cy="914400"/>
      <wp:docPr id="9" name="Group with unsupported child"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
        <wpg:wgp>
          <wpg:cNvPr id="1" name="Synthetic group"/>
          <wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>
            <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></wpg:grpSpPr>
          <synthetic:shape id="11" kept="true">
            <a:cNvPr id="12" name="Unknown child">
              <a:hlinkHover r:id="rId2" tooltip="Child hover"/>
            </a:cNvPr>
            <synthetic:data value="kept"/>
          </synthetic:shape>
        </wpg:wgp>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p></w:body>
</w:document>`;

function syntheticSource() {
  return zipSync({
    "[Content_Types].xml": ENCODER.encode(CONTENT_TYPES_XML),
    "word/document.xml": ENCODER.encode(DOCUMENT_XML),
    "word/_rels/document.xml.rels": ENCODER.encode(RELS_XML),
  });
}

function unsupportedChild(options: DocumentOptions) {
  const child = options.sections?.[0]?.children?.[0];
  if (!child || !("paragraph" in child)) throw new Error("synthetic fixture has no paragraph");
  const paragraph = typeof child.paragraph === "string" ? undefined : child.paragraph;
  if (!paragraph) throw new Error("synthetic fixture has no paragraph options");
  return paragraph?.children?.find(
    (
      candidate,
    ): candidate is Extract<Exclude<ParagraphChild, string>, { unsupportedDrawing: unknown }> =>
      typeof candidate === "object" && "unsupportedDrawing" in candidate,
  );
}

function groupSource() {
  return zipSync({
    "[Content_Types].xml": ENCODER.encode(CONTENT_TYPES_XML),
    "word/document.xml": ENCODER.encode(GROUP_XML),
    "word/_rels/document.xml.rels": ENCODER.encode(RELS_XML),
  });
}

function alternateContentSource() {
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return zipSync({
    "[Content_Types].xml": ENCODER.encode(CONTENT_TYPES_XML),
    "word/document.xml": ENCODER.encode(ALTERNATE_CONTENT_XML),
    "word/_rels/document.xml.rels": ENCODER.encode(ALTERNATE_CONTENT_RELS_XML),
    "word/media/image1.png": png,
  });
}

function groupDrawing(options: DocumentOptions) {
  const child = options.sections?.[0]?.children?.[0];
  if (!child || !("paragraph" in child)) throw new Error("group fixture has no paragraph");
  if (typeof child.paragraph === "string") throw new Error("group fixture has no runs");
  const drawing = child.paragraph.children?.find(
    (candidate): candidate is Extract<Exclude<ParagraphChild, string>, { wpgGroup: unknown }> =>
      typeof candidate === "object" && "wpgGroup" in candidate,
  );
  if (!drawing) throw new Error("group fixture has no wpgGroup");
  return drawing;
}

describe("unsupported DrawingML fallback", () => {
  it("canonicalizes the payload and preserves its hover relationship", async () => {
    const first = parseDocumentSync(syntheticSource()) as DocumentOptions;
    const drawing = unsupportedChild(first);
    expect(drawing?.unsupportedDrawing.graphicDataUri).toBe(
      "https://example.invalid/synthetic-drawing",
    );
    expect(drawing?.unsupportedDrawing.additionRsid).toBe("00AA0001");
    expect(drawing?.unsupportedDrawing.runProperties).toMatchObject({ bold: true });
    expect(drawing?.unsupportedDrawing.relationships).toEqual([
      {
        sourceRelationshipId: "rId2",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.invalid/hover",
        targetMode: "External",
      },
    ]);

    const generated = await generateDocument(first, { type: "uint8array" });
    const files = unzipSync(generated);
    const documentXml = DECODER.decode(files["word/document.xml"]!);
    const relsXml = DECODER.decode(files["word/_rels/document.xml.rels"]!);
    expect(documentXml).toContain('r:id="rId2"');
    expect(relsXml).toContain('Id="rId2"');
    expect(relsXml).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"',
    );
    expect(relsXml).toContain('Target="https://example.invalid/hover" TargetMode="External"');

    const second = parseDocumentSync(generated) as DocumentOptions;
    const reparsed = unsupportedChild(second);
    expect(reparsed?.unsupportedDrawing.relationships).toEqual(
      drawing?.unsupportedDrawing.relationships,
    );
    expect(reparsed?.unsupportedDrawing.graphicData).toEqual(
      drawing?.unsupportedDrawing.graphicData,
    );
    expect("rawXml" in (reparsed ?? {})).toBe(false);
  });

  it("keeps an unrecognized group child and its hover relationship", async () => {
    const first = parseDocumentSync(groupSource()) as DocumentOptions;
    const child = groupDrawing(first).wpgGroup.children[0];
    if (!child) throw new Error("unknown group child was dropped");
    expect(child).toMatchObject({
      type: "unsupported",
      element: { name: "synthetic:shape", attributes: { id: 11, kept: "true" } },
      relationships: [
        {
          sourceRelationshipId: "rId2",
          relationshipType:
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
          target: "https://example.invalid/hover",
          targetMode: "External",
        },
      ],
    });

    const second = parseDocumentSync(await generateDocument(first, { type: "uint8array" }));
    const regenerated = groupDrawing(second as DocumentOptions).wpgGroup.children[0];
    expect(regenerated).toMatchObject(child);
  });
});

describe("unsupported DrawingML alternate content", () => {
  it("preserves the mc:AlternateContent wrapper, run properties, and fallback media", async () => {
    const first = parseDocumentSync(alternateContentSource()) as DocumentOptions;
    const child = first.sections?.[0]?.children?.[0];
    if (!child || !("paragraph" in child)) throw new Error("fixture has no paragraph");
    if (typeof child.paragraph === "string") throw new Error("fixture has no runs");
    const drawing = child.paragraph.children?.find(
      (
        candidate,
      ): candidate is Extract<Exclude<ParagraphChild, string>, { unsupportedDrawing: unknown }> =>
        typeof candidate === "object" && "unsupportedDrawing" in candidate,
    );
    if (!drawing) throw new Error("alternate content fixture has no unsupportedDrawing");
    expect(drawing.unsupportedDrawing.mcChoiceRequires).toBe("cx1");
    expect(drawing.unsupportedDrawing.runProperties).toMatchObject({ noProof: true });
    expect(drawing.unsupportedDrawing.vmlFallback).toContain("<mc:Fallback>");

    const generated = await generateDocument(first, { type: "uint8array" });
    const files = unzipSync(generated);
    const documentXml = DECODER.decode(files["word/document.xml"]!);
    expect(documentXml).toContain("<w:rPr><w:noProof/></w:rPr>");
    expect(documentXml).toContain('<mc:Choice Requires="cx1">');
    expect(documentXml).toContain("<mc:Fallback>");

    const second = parseDocumentSync(generated) as DocumentOptions;
    const secondChild = second.sections?.[0]?.children?.[0];
    if (!secondChild || !("paragraph" in secondChild)) throw new Error("round-trip lost paragraph");
    if (typeof secondChild.paragraph === "string") throw new Error("round-trip lost runs");
    const reparsed = secondChild.paragraph.children?.find(
      (
        candidate,
      ): candidate is Extract<Exclude<ParagraphChild, string>, { unsupportedDrawing: unknown }> =>
        typeof candidate === "object" && "unsupportedDrawing" in candidate,
    );
    expect(reparsed?.unsupportedDrawing.mcChoiceRequires).toBe("cx1");
    expect(reparsed?.unsupportedDrawing.runProperties).toMatchObject({ noProof: true });
    expect(reparsed?.unsupportedDrawing.vmlFallback).toContain("<mc:Fallback>");
  });
});
