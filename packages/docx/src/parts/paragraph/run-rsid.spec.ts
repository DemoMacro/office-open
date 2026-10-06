import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph, stringifyParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

const readCtx = {} as unknown as DocxReadContext;
const writeCtx = {} as never;

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';

function parseParagraphXml(xml: string) {
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed paragraph has no root element");
  return parseParagraph(el, readCtx);
}

function firstChild(opts: ReturnType<typeof parseParagraphXml>): Record<string, any> {
  const child = opts.children?.[0];
  if (child === undefined) throw new Error("paragraph has no run child");
  return child as Record<string, any>;
}

describe("run rsid round-trip", () => {
  it("preserves w:rsidR and w:rsidRPr on text runs", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD"><w:t>hi</w:t></w:r></w:p>`,
    );
    expect(firstChild(opts)).toMatchObject({
      text: "hi",
      additionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">',
    );
  });

  it("preserves w:rsidR and w:rsidRPr on every plain complex-field run", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}>` +
        '<w:r w:rsidR="00000001" w:rsidRPr="00000002"><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r w:rsidR="00000003" w:rsidRPr="00000004"><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r w:rsidR="00000005" w:rsidRPr="00000006"><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r w:rsidR="00000007" w:rsidRPr="00000008"><w:t>1</w:t></w:r>' +
        '<w:r w:rsidR="00000009" w:rsidRPr="0000000A"><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:p>",
    );
    expect(firstChild(opts).complexField).toMatchObject({
      additionRsid: "00000001",
      runPropertiesRsid: "00000002",
      instructionAdditionRsid: "00000003",
      instructionRunPropertiesRsid: "00000004",
      separatorAdditionRsid: "00000005",
      separatorRunPropertiesRsid: "00000006",
      resultAdditionRsid: "00000007",
      resultRunPropertiesRsid: "00000008",
      endAdditionRsid: "00000009",
      endRunPropertiesRsid: "0000000A",
    });
    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml).toContain('w:rsidR="00000001" w:rsidRPr="00000002"');
    expect(xml).toContain('w:rsidR="00000003" w:rsidRPr="00000004"');
    expect(xml).toContain('w:rsidR="00000005" w:rsidRPr="00000006"');
    expect(xml).toContain('w:rsidR="00000007" w:rsidRPr="00000008"');
    expect(xml).toContain('w:rsidR="00000009" w:rsidRPr="0000000A"');
  });

  it("preserves run identity on page, column, textWrapping, and tab runs", () => {
    for (const [inner, expected] of [
      ['<w:br w:type="page"/>', '<w:br w:type="page"/>'],
      ['<w:br w:type="column"/>', '<w:br w:type="column"/>'],
      ['<w:br w:type="textWrapping"/>', '<w:br w:type="textWrapping"/>'],
      ["<w:tab/>", "<w:tab/>"],
    ] as const) {
      const opts = parseParagraphXml(
        `<w:p ${NS}><w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">${inner}</w:r></w:p>`,
      );
      expect(opts.children?.[0]).toMatchObject({
        additionRsid: "00112233",
        runPropertiesRsid: "AABBCCDD",
      });
      expect(stringifyParagraph(opts, writeCtx)).toContain(
        `<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">${expected}</w:r>`,
      );
    }
  });

  it("preserves w:rsidR and w:rsidRPr on hyperlink text runs", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:hyperlink w:anchor="target">` +
        '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD"><w:t>link</w:t></w:r>' +
        "</w:hyperlink></w:p>",
    );
    const hyperlink = firstChild(opts).hyperlink;
    expect(hyperlink.children[0]).toMatchObject({
      additionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">',
    );
  });

  it("does not copy result-run rsids onto the field separator", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:hyperlink w:anchor="target">` +
        '<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r w:rsidR="00112233"><w:rPr><w:webHidden/></w:rPr><w:instrText> PAGEREF </w:instrText></w:r>' +
        '<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r w:rsidR="00445566"><w:rPr><w:webHidden/></w:rPr><w:t>7</w:t></w:r>' +
        '<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:hyperlink></w:p>",
    );
    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml).toContain(
      '<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>',
    );
    expect(xml).toContain(
      '<w:r w:rsidR="00445566"><w:rPr><w:webHidden/></w:rPr><w:t>7</w:t></w:r>',
    );
  });

  it("preserves run identity on pure reference runs", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:commentRangeStart w:id="7"/>` +
        '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">' +
        '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
        '<w:commentReference w:id="7"/></w:r>' +
        "</w:p>",
    );
    const reference = opts.children?.find(
      (child) => child !== null && typeof child === "object" && "commentReference" in child,
    ) as Record<string, unknown>;
    expect(reference).toMatchObject({
      commentReference: 7,
      additionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
      properties: { style: "CommentReference" },
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">',
    );
  });

  it("preserves run identity around a VML picture", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD" w:rsidDel="99887766">` +
        '<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:rect/></w:pict>' +
        "</w:r></w:p>",
    );
    expect(firstChild(opts).pict).toMatchObject({
      runAdditionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
      runDeletionRsid: "99887766",
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD" w:rsidDel="99887766">',
    );
  });

  it("preserves run identity on symbol runs", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">` +
        '<w:sym w:font="Wingdings" w:char="F0E0"/></w:r></w:p>',
    );
    expect(firstChild(opts).symbolRun).toMatchObject({
      additionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">',
    );
  });

  it("preserves non-plain simple-field result runs structurally", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:fldSimple w:instr=" PAGE ">` +
        '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">' +
        "<w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>" +
        "</w:fldSimple></w:p>",
    );
    const field = firstChild(opts).simpleField;
    expect(field.cachedValue).toBe("1");
    expect(field.cachedRuns).toMatchObject([
      {
        text: "1",
        additionRsid: "00112233",
        runPropertiesRsid: "AABBCCDD",
        noProof: true,
      },
    ]);
    const xml = stringifyParagraph(opts, writeCtx);
    expect(xml).toContain('<w:fldSimple w:instr=" PAGE ">');
    expect(xml).toContain('<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD"><w:rPr>');
  });

  it("round-trips a plain simple-field instruction-text result", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:fldSimple w:instr=" PAGE ">` +
        '<w:r><w:instrText xml:space="preserve">A</w:instrText></w:r>' +
        "</w:fldSimple></w:p>",
    );
    expect(firstChild(opts).simpleField).toMatchObject({
      cachedInstructionText: "A",
      cachedInstructionTextPreserveSpace: true,
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      '<w:r><w:instrText xml:space="preserve">A</w:instrText></w:r>',
    );
  });

  it("preserves run properties on a simple-field instruction-text result", () => {
    const opts = parseParagraphXml(
      `<w:p ${NS}><w:fldSimple w:instr=" SEQ Appendix ">` +
        "<w:r><w:rPr><w:noProof/></w:rPr><w:instrText>A</w:instrText></w:r>" +
        "</w:fldSimple></w:p>",
    );
    expect(firstChild(opts).simpleField).toMatchObject({
      cachedInstructionText: "A",
      cachedInstructionRPrXml: "<w:rPr><w:noProof/></w:rPr>",
    });
    expect(stringifyParagraph(opts, writeCtx)).toContain(
      "<w:r><w:rPr><w:noProof/></w:rPr><w:instrText>A</w:instrText></w:r>",
    );
  });

  it("preserves w:rsidR and w:rsidRPr on drawing runs", () => {
    const chartElement = parseXml(
      `<root ${NS}><c:chartSpace><c:chart><c:plotArea><c:barChart>` +
        '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit>' +
        '<c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
        "</c:barChart></c:plotArea></c:chart></c:chartSpace></root>",
    ).elements?.[0]?.elements?.[0];
    const chartReadContext = {
      docx: {
        partRefs: { charts: new Map([["rId1", "word/charts/chart1.xml"]]) },
        doc: {
          get: (path: string) => (path === "word/charts/chart1.xml" ? chartElement : undefined),
        },
      },
    } as unknown as DocxReadContext;
    const paragraphElement = parseXml(
      `<w:p ${NS}><w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD"><w:drawing>` +
        '<wp:inline><wp:extent cx="5486400" cy="3200400"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
        '<c:chart r:id="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    ).elements?.[0];
    if (!paragraphElement) throw new Error("parsed paragraph has no root element");
    const opts = parseParagraph(paragraphElement, chartReadContext);
    expect(firstChild(opts).chart).toMatchObject({
      additionRsid: "00112233",
      runPropertiesRsid: "AABBCCDD",
    });
    const chartWriteContext = {
      file: { charts: { addChart: () => undefined } },
      fileData: {},
      reproducible: { nextDrawingId: () => 1 },
      stringifyChild: () => "",
      viewWrapper: { relationships: { hasId: () => false, add: () => 1 } },
    } as never;
    expect(stringifyParagraph(opts, chartWriteContext)).toContain(
      '<w:r w:rsidR="00112233" w:rsidRPr="AABBCCDD">',
    );
  });
});
