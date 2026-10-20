import { parse } from "@office-open/xml";
import { describe, expect, it } from "vitest";

import type { DocxReadContext } from "../../context";
import type { BodyContext } from "../../context";
import { settingsDesc } from "../settings/descriptor";
import {
  parseMathChildren,
  parseMathParagraph,
  stringifyMathInput,
  stringifyMathParagraph,
} from "./math/stringify";

const NAMESPACES =
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseElement(xml: string) {
  const root = parse(xml).elements?.[0];
  if (!root) throw new Error("math XML did not parse");
  return root;
}

describe("math run text", () => {
  it("round-trips preserved spaces", () => {
    const xml =
      '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
      '<m:r><m:t xml:space="preserve"> x </m:t></m:r></m:oMath>';
    const root = parse(xml).elements?.[0];
    if (!root) throw new Error("math XML did not parse");
    const children = parseMathChildren(root);
    expect(children).toEqual([{ text: " x ", preserveSpace: true }]);
    expect(stringifyMathInput(children[0]!)).toContain('<m:t xml:space="preserve"> x </m:t>');
  });

  it("round-trips a Word comment reference as mixed math-run content", () => {
    const root = parseElement(
      `<m:oMath ${NAMESPACES}><m:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
        '<w:commentReference w:id="1"/></m:r></m:oMath>',
    );
    const children = parseMathChildren(root);
    expect(children).toEqual([
      {
        children: [{ commentReference: 1 }],
        runProperties: { style: "CommentReference" },
      },
    ]);
    expect(stringifyMathInput(children[0]!)).toContain('<w:commentReference w:id="1"/>');
  });

  it("preserves custom n-ary operators", () => {
    const root = parseElement(
      `<m:oMath ${NAMESPACES}><m:nary><m:naryPr><m:chr m:val="∏"/>` +
        '<m:limLoc m:val="undOvr"/><m:grow m:val="off"/></m:naryPr>' +
        "<m:sub><m:r><m:t>i</m:t></m:r></m:sub>" +
        "<m:sup><m:r><m:t>n</m:t></m:r></m:sup>" +
        "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>",
    );
    const parsed = parseMathChildren(root);
    expect(parsed).toEqual([
      {
        nary: {
          character: "∏",
          children: ["x"],
          subScript: ["i"],
          superScript: ["n"],
          properties: { limitLocation: "undOvr", grow: false, growRaw: "off" },
        },
      },
    ]);
    const xml = stringifyMathInput(parsed[0]!);
    expect(xml).toContain('<m:chr m:val="∏"/>');
    expect(parseMathChildren(parseElement(`<m:oMath ${NAMESPACES}>${xml}</m:oMath>`))).toEqual(
      parsed,
    );
  });

  it("preserves n-ary toggle absence and source spelling", () => {
    const root = parseElement(
      `<m:oMath ${NAMESPACES}><m:nary><m:naryPr><m:chr m:val="∑"/>` +
        '<m:grow m:val="1"/></m:naryPr><m:sub/><m:sup>' +
        "<m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>",
    );
    const parsed = parseMathChildren(root);
    expect(parsed).toEqual([
      {
        sum: {
          children: ["x"],
          superScript: ["n"],
          properties: { grow: true, growRaw: "1" },
        },
      },
    ]);
    const xml = stringifyMathInput(parsed[0]!);
    expect(xml).toContain('<m:grow m:val="1"/>');
    expect(xml).not.toContain("m:subHide");
    expect(xml).not.toContain("m:supHide");
  });

  it("round-trips multiple display equations", () => {
    const root = parseElement(
      `<m:oMathPara ${NAMESPACES}><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>` +
        "<m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>" +
        "<m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara>",
    );
    const parsed = parseMathParagraph(root);
    expect(parsed).toEqual({
      equations: [["a"], ["b"]],
      justification: "center",
    });
    const xml = stringifyMathParagraph(parsed.equations, parsed.justification);
    expect(parseMathParagraph(parseElement(xml))).toEqual(parsed);
  });

  it("round-trips an empty argument property element", () => {
    const root = parseElement(
      `<m:oMath ${NAMESPACES}><m:sSup><m:e><m:argPr/><m:r><m:t>x</m:t></m:r></m:e>` +
        "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>",
    );
    const parsed = parseMathChildren(root);
    expect(parsed[0]).toMatchObject({
      superScript: { baseArgumentProperties: {}, children: ["x"] },
    });
    expect(stringifyMathInput(parsed[0]!)).toContain(
      "<m:e><m:argPr/><m:r><m:t>x</m:t></m:r></m:e>",
    );
  });

  it("round-trips inserted and deleted control properties", () => {
    const inserted = parseMathChildren(
      parseElement(
        `<m:oMath ${NAMESPACES}><m:f><m:fPr><m:ctrlPr><w:ins w:id="7" w:author="Alice">` +
          "<w:rPr><w:b/></w:rPr></w:ins></m:ctrlPr></m:fPr>" +
          "<m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>",
      ),
    )[0];
    expect(inserted).toMatchObject({
      fraction: {
        controlProperties: { insertion: { id: 7, author: "Alice", runProperties: { bold: true } } },
      },
    });
    expect(stringifyMathInput(inserted!)).toContain(
      '<w:ins w:id="7" w:author="Alice"><w:rPr><w:b/></w:rPr></w:ins>',
    );

    const deleted = parseMathChildren(
      parseElement(
        `<m:oMath ${NAMESPACES}><m:f><m:fPr><m:ctrlPr><w:del w:id="8" w:author="Bob">` +
          "<w:rPr><w:i/></w:rPr></w:del></m:ctrlPr></m:fPr>" +
          "<m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>",
      ),
    )[0];
    expect(deleted).toMatchObject({
      fraction: {
        controlProperties: {
          deletion: { id: 8, author: "Bob", runProperties: { italic: true } },
        },
      },
    });
    expect(stringifyMathInput(deleted!)).toContain(
      '<w:del w:id="8" w:author="Bob"><w:rPr><w:i/></w:rPr></w:del>',
    );
  });

  it("round-trips mixed Word run content", () => {
    const parsed = parseMathChildren(
      parseElement(
        `<m:oMath ${NAMESPACES}><m:r><m:t>a</m:t><w:tab/><w:br/><m:t>b</m:t></m:r></m:oMath>`,
      ),
    );
    expect(parsed).toEqual([{ children: ["a", { tab: true }, { break: 1 }, "b"] }]);
    expect(stringifyMathInput(parsed[0]!)).toBe(
      "<m:r><m:t>a</m:t><w:tab/><w:br/><m:t>b</m:t></m:r>",
    );
  });

  it("round-trips mixed Word run drawing content", () => {
    const readContext = {
      currentPart: "word/document.xml",
      consumedPartPaths: new Set<string>(),
      docx: {
        doc: new Map(),
        partRefs: {
          partMedia: new Map(),
          headers: new Map(),
          footers: new Map(),
          media: new Map(),
          charts: new Map(),
          diagramData: new Map(),
          afChunks: new Map(),
          subDocs: new Map(),
          hyperlinks: new Map(),
          partEmbeddingTypes: new Map(),
          partExternalImages: new Map(),
        },
      },
    } as unknown as DocxReadContext;
    const parsed = parseMathChildren(
      parseElement(
        `<m:oMath ${NAMESPACES} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
          "<m:r><m:t>d</m:t><w:tab/><w:br/>" +
          '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
          '<a:graphic><a:graphicData uri="https://example.invalid/math-drawing"/></a:graphic>' +
          "</wp:inline></w:drawing></m:r></m:oMath>",
      ),
      readContext,
    );
    expect(parsed[0]).toMatchObject({
      children: [
        "d",
        { tab: true },
        { break: 1 },
        {
          unsupportedDrawing: {
            anchorMode: "inline",
            graphicDataUri: "https://example.invalid/math-drawing",
          },
        },
      ],
    });
    const writeContext = { fileData: {} } as unknown as BodyContext;
    const xml = `<m:oMath ${NAMESPACES} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${stringifyMathInput(parsed[0]!, writeContext)}</m:oMath>`;
    expect(parseMathChildren(parseElement(xml), readContext)).toEqual(parsed);
  });

  it("keeps explicit false and zero math values", () => {
    const parsed = parseMathChildren(
      parseElement(
        `<m:oMath ${NAMESPACES}><m:borderBox><m:borderBoxPr><m:hideTop m:val="off"/>` +
          "</m:borderBoxPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:borderBox></m:oMath>",
      ),
    );
    expect(parsed[0]).toMatchObject({ borderBox: { properties: { hideTop: false } } });
    expect(stringifyMathInput(parsed[0]!)).toContain('<m:hideTop m:val="off"/>');
  });

  it("keeps an explicitly disabled radical degree", () => {
    const parsed = parseMathChildren(
      parseElement(
        `<m:oMath ${NAMESPACES}><m:rad><m:radPr><m:degHide m:val="off"/></m:radPr>` +
          "<m:deg><m:r><m:t>n</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>",
      ),
    );
    expect(parsed[0]).toMatchObject({ radical: { properties: { hideDegree: false } } });
    expect(stringifyMathInput(parsed[0]!)).toContain('<m:degHide m:val="off"/>');
  });

  it("round-trips math property measurements", () => {
    const source = parseElement(
      `<w:settings ${NAMESPACES}><m:mathPr><m:lMargin m:val="720"/><m:preSp m:val="10pt"/></m:mathPr></w:settings>`,
    );
    const parsed = settingsDesc.parse(source, {} as never);
    expect(parsed).toMatchObject({
      mathProperties: { leftMargin: 720, preSpacing: "10pt" },
    });
    const xml = settingsDesc.stringify(
      { mathProperties: { leftMargin: "0.5in", preSpacing: "10pt" } },
      {} as never,
    );
    expect(xml).toContain('<m:lMargin m:val="0.5in"/>');
    expect(xml).toContain('<m:preSp m:val="10pt"/>');
    expect(settingsDesc.parse(parseElement(xml!), {} as never)).toMatchObject({
      mathProperties: { leftMargin: "0.5in", preSpacing: "10pt" },
    });
  });
});
