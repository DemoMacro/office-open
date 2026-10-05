import { unzipSync } from "../../../packages/core/dist/index.mjs";
import { generateDocument, parseDocument } from "../../../packages/docx/dist/index.mjs";
import { generatePresentation, parsePresentation } from "../../../packages/pptx/dist/index.mjs";
import { generateWorkbook, parseWorkbook } from "../../../packages/xlsx/dist/index.mjs";
import {
  ENCODER,
  assert,
  assertEqual,
  auditSyntheticOptions,
  partText,
  syntheticZip,
} from "./support";

const PRINTER_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";

export async function docxPrinterSettings(): Promise<void> {
  const part = "word/document.xml";
  const document = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:printerSettings r:id="rId7"/></w:sectPr></w:body></w:document>`;
  const relationships = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="${PRINTER_RELATIONSHIP_TYPE}" Target="printerSettings/settings.bin"/></Relationships>`;
  const source = syntheticZip({
    "[Content_Types].xml": ENCODER.encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": ENCODER.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": ENCODER.encode(document),
    "word/_rels/document.xml.rels": ENCODER.encode(relationships),
    "word/printerSettings/settings.bin": new Uint8Array([1, 2, 3, 4]),
  });
  const options = await parseDocument(source);
  assertEqual(
    options.sections[0]?.properties?.printerSettingsPath,
    "word/printerSettings/settings.bin",
    part,
    "printer settings path projection",
  );
  auditSyntheticOptions(options, part);
  const output = await generateDocument(options, { type: "uint8array" });
  assertEqual(
    unzipSync(output)["word/printerSettings/settings.bin"]?.join(","),
    "1,2,3,4",
    part,
    "printer settings binary projection",
  );
  assert(
    partText(output, "word/_rels/document.xml.rels").includes(PRINTER_RELATIONSHIP_TYPE),
    part,
    "printer settings relationship projection",
  );
}

export async function pptxIndefiniteTiming(): Promise<void> {
  const part = "ppt/slides/slide1.xml";
  const seed = await generatePresentation(
    {
      slides: [
        {
          children: [{ shape: { id: 2, x: 1, y: 2, width: 3, height: 4 } }],
          animations: [{ shapeId: 2, type: "fade", trigger: "onClick", duration: "indefinite" }],
        },
      ],
    },
    { type: "uint8array" },
  );
  const seedSlide = partText(seed, "ppt/slides/slide1.xml");
  const timingStart = seedSlide.indexOf("<p:timing>");
  const timingEnd = seedSlide.lastIndexOf("</p:timing>");
  assert(timingStart >= 0 && timingEnd > timingStart, part, "seed timing is missing");
  const timing = seedSlide.slice(timingStart, timingEnd + "</p:timing>".length);
  const slide = `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></p:spTree></p:cSld>${timing}</p:sld>`;
  const presentation = `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const source = syntheticZip({
    "[Content_Types].xml": ENCODER.encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
    ),
    "_rels/.rels": ENCODER.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    ),
    "ppt/presentation.xml": ENCODER.encode(presentation),
    "ppt/_rels/presentation.xml.rels": ENCODER.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`,
    ),
    "ppt/slides/slide1.xml": ENCODER.encode(slide),
  });
  const options = await parsePresentation(source);
  assertEqual(
    options.slides?.[0]?.animations?.[0]?.duration,
    "indefinite",
    part,
    "duration projection",
  );
  auditSyntheticOptions(options, part);
  const output = await generatePresentation(options, { type: "uint8array" });
  assert(
    partText(output, "ppt/slides/slide1.xml").includes('dur="indefinite"'),
    part,
    "generated indefinite timing",
  );
}

export async function xlsxPivotCalculatedItems(): Promise<void> {
  const part = "xl/pivotCache/pivotCacheDefinition1.xml";
  const workbook = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><pivotCaches><pivotCache cacheId="1" r:id="rId2"/></pivotCaches></workbook>`;
  const worksheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>B</t></is></c><c r="B3"><v>2</v></c></row></sheetData></worksheet>`;
  const definition = `<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="2"><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Data"/></cacheSource><cacheFields count="2"><cacheField name="Name" numFmtId="0"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField><cacheField name="Value" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="2" count="2"><n v="1"/><n v="2"/></sharedItems></cacheField></cacheFields><calculatedItems count="1"><calculatedItem field="0" formula="A*2"><pivotArea field="0"><references count="1"><reference field="0" count="1"><x v="1"/></reference></references></pivotArea></calculatedItem></calculatedItems></pivotCacheDefinition>`;
  const records = `<?xml version="1.0"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2"><r><x v="0"/><n v="1"/></r><r><x v="1"/><n v="2"/></r></pivotCacheRecords>`;
  const source = syntheticZip({
    "[Content_Types].xml": ENCODER.encode(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/></Types>`,
    ),
    "_rels/.rels": ENCODER.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    "xl/workbook.xml": ENCODER.encode(workbook),
    "xl/_rels/workbook.xml.rels": ENCODER.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": ENCODER.encode(worksheet),
    "xl/pivotCache/pivotCacheDefinition1.xml": ENCODER.encode(definition),
    "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels": ENCODER.encode(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>`,
    ),
    "xl/pivotCache/pivotCacheRecords1.xml": ENCODER.encode(records),
  });
  const options = await parseWorkbook(source);
  assertEqual(
    options.pivotCaches?.[0]?.definition?.calculatedItems?.[0]?.formula,
    "A*2",
    part,
    "calculated item formula projection",
  );
  assertEqual(
    options.pivotCaches?.[0]?.definition?.calculatedItems?.[0]?.pivotArea?.references?.[0]?.x?.[0],
    1,
    part,
    "calculated item pivot area projection",
  );
  auditSyntheticOptions(options, part);
  const output = await generateWorkbook(options, { type: "uint8array" });
  const generated = partText(output, "xl/pivotCache/pivotCacheDefinition1.xml");
  assert(generated.includes('formula="A*2"'), part, "generated calculated item");
  assert(generated.includes('<x v="1"/>'), part, "generated calculated item reference");
}
