import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraphProperties } from "../../body";
import { stringifyParagraphProperties } from "./stringify";

describe("stringifyParagraphProperties single-writer channels", () => {
  it("preserves an empty source paragraph border", () => {
    const element = parseXml('<w:pPr xmlns:w="urn:w"><w:pBdr/></w:pPr>').elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    expect(parsed.border).toEqual({});
    expect(stringifyParagraphProperties(parsed).xml).toContain("<w:pBdr>");
  });

  it("preserves noncanonical source border and shading lexical forms", () => {
    const xml =
      `<w:pPr xmlns:w="urn:w">` +
      `<w:pBdr><w:bottom w:val="single" w:color="darkblue"/></w:pBdr>` +
      `<w:shd w:val="clear" w:color="auto" w:themeTint="0" w:themeShade="00" w:fill="D9D9D9" ` +
      `w:themeFillTint="0" w:themeFillShade="d9"/></w:pPr>`;
    const element = parseXml(xml).elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    expect(parsed.border?.bottom).toMatchObject({ color: "darkblue", colorRaw: "darkblue" });
    expect(parsed.shading).toMatchObject({
      color: "auto",
      themeTintRaw: "0",
      themeShade: "00",
      fill: "D9D9D9",
      themeFillTintRaw: "0",
      themeFillShade: "d9",
    });
    const { xml: output } = stringifyParagraphProperties(parsed);
    expect(output).toContain('w:color="darkblue"');
    expect(output).toContain('w:themeTint="0"');
    expect(output).toContain('w:themeFillTint="0"');
  });

  it("preserves legacy paragraph inside borders", () => {
    const element = parseXml(
      '<w:pPr xmlns:w="urn:w"><w:pBdr>' +
        '<w:insideH w:val="single" w:sz="6"/>' +
        '<w:insideV w:val="nil"/></w:pBdr></w:pPr>',
    ).elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    expect(parsed.border?.insideHorizontal).toMatchObject({ style: "single", size: 6 });
    expect(parsed.border?.insideVertical).toMatchObject({ style: "nil" });
    const { xml } = stringifyParagraphProperties(parsed);
    expect(xml).toContain('<w:insideH w:val="single" w:sz="6"/>');
    expect(xml).toContain('<w:insideV w:val="nil"/>');
  });

  it("writes paragraph borders in XSD order", () => {
    const element = parseXml(
      '<w:pPr xmlns:w="urn:w"><w:pBdr>' +
        '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/>' +
        "</w:pBdr></w:pPr>",
    ).elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    const { xml } = stringifyParagraphProperties(parsed);
    expect(xml).toContain(
      '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/>',
    );
  });

  it("preserves an explicit empty spacing element", () => {
    const element = parseXml('<w:pPr xmlns:w="urn:w"><w:spacing/></w:pPr>').elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    expect(parsed.spacing).toEqual({});
    expect(stringifyParagraphProperties(parsed).xml).toContain("<w:spacing/>");
  });

  it("preserves spacing measure lexical forms", () => {
    const element = parseXml(
      '<w:pPr xmlns:w="urn:w">' +
        '<w:spacing w:before="12.0pt" w:after="6.50pt" w:line="13.80pt" w:lineRule="auto"/>' +
        "</w:pPr>",
    ).elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    expect(parsed.spacing).toMatchObject({
      before: "12.0pt",
      after: "6.50pt",
      line: "13.80pt",
    });
    expect(stringifyParagraphProperties(parsed).xml).toContain(
      '<w:spacing w:after="6.50pt" w:before="12.0pt" w:line="13.80pt" w:lineRule="auto"/>',
    );
  });

  it("preserves a source-omitted shading value without fresh clear", () => {
    const element = parseXml('<w:pPr xmlns:w="urn:w"><w:shd w:fill="CCCCCC"/></w:pPr>')
      .elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    const parsed = parseParagraphProperties(element, {} as never);
    expect(parsed.shading).toMatchObject({ type: null, fill: "CCCCCC" });
    const roundTrip = stringifyParagraphProperties(parsed);
    expect(roundTrip.xml).toContain('<w:shd w:fill="CCCCCC"/>');
    expect(stringifyParagraphProperties({ shading: { fill: "CCCCCC" } }).xml).toContain(
      '<w:shd w:val="clear" w:fill="CCCCCC"/>',
    );
  });

  it("preserves an explicit contextual-spacing spelling", () => {
    const { xml } = stringifyParagraphProperties({
      contextualSpacing: false,
      contextualSpacingRaw: "0",
    });
    expect(xml).toContain('<w:contextualSpacing w:val="0"/>');
    const element = parseXml('<w:pPr xmlns:w="urn:w"><w:contextualSpacing w:val="0"/></w:pPr>')
      .elements?.[0];
    if (!element) throw new Error("parsed pPr has no root element");
    expect(parseParagraphProperties(element, {} as never)).toMatchObject({
      contextualSpacing: false,
      contextualSpacingRaw: "0",
    });
  });

  it("emits w15:collapsed with explicit on and off states", () => {
    const on = stringifyParagraphProperties({ collapsed: true });
    expect(on.xml).toContain("<w15:collapsed/>");
    const off = stringifyParagraphProperties({ collapsed: false });
    expect(off.xml).toContain('<w15:collapsed w:val="off"/>');
  });

  it("emits one w:pStyle when heading and bullet are combined", () => {
    // CT_PPrBase allows exactly one w:pStyle; heading wins over the
    // ListParagraph sugar.
    const { xml } = stringifyParagraphProperties({ heading: "Heading1", bullet: { level: 0 } });
    expect(xml?.match(/<w:pStyle/g)).toHaveLength(1);
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>');
  });

  it("emits one w:pStyle when style and bullet are combined", () => {
    const { xml } = stringifyParagraphProperties({ style: "MyList", bullet: { level: 0 } });
    expect(xml?.match(/<w:pStyle/g)).toHaveLength(1);
    expect(xml).toContain('<w:pStyle w:val="MyList"/>');
    // The bullet numPr still applies under the named style.
    expect(xml).toContain('<w:numId w:val="1"/>');
  });

  it("lets an explicit numbering win over the bullet sugar", () => {
    const { xml } = stringifyParagraphProperties({
      bullet: { level: 2 },
      numbering: { reference: "ref", level: 1 },
    });
    expect(xml?.match(/<w:numPr/g)).toHaveLength(1);
    expect(xml).toContain('<w:numId w:val="{ref-0}"/>');
  });

  it("lets numbering false (remove) win over the bullet sugar", () => {
    const { xml } = stringifyParagraphProperties({ bullet: { level: 0 }, numbering: false });
    expect(xml?.match(/<w:numPr/g)).toHaveLength(1);
    expect(xml).toContain('<w:numId w:val="0"/>');
  });

  it("omits w:ilvl when the numbering carries no level (numId-only numPr)", () => {
    const { xml } = stringifyParagraphProperties({
      numbering: { reference: "ref", autoStyle: false },
    });
    expect(xml).toContain('<w:numId w:val="{ref-0}"/>');
    expect(xml).not.toContain("<w:ilvl");
  });

  it("preserves unknown source numbering IDs without fabricating a bullet", () => {
    const parsed = parseParagraphProperties(
      parseXml('<w:pPr xmlns:w="urn:w"><w:numPr><w:numId w:val="32"/></w:numPr></w:pPr>')
        .elements![0]!,
      {} as never,
    );
    expect(parsed.numbering).toEqual({
      reference: "list_32",
      sourceNumId: 32,
      autoStyle: false,
    });
    const { xml } = stringifyParagraphProperties(parsed);
    expect(xml).toContain('<w:numPr><w:numId w:val="32"/></w:numPr>');
  });

  it("merges thematicBreak into an explicit border as the bottom edge", () => {
    const { xml } = stringifyParagraphProperties({
      thematicBreak: true,
      border: { top: { color: "auto", size: 4, space: 1, style: "single" } },
    });
    expect(xml?.match(/<w:pBdr/g)).toHaveLength(1);
    expect(xml).toContain("<w:top");
    expect(xml).toContain("<w:bottom");
  });

  it("keeps an explicit border bottom over thematicBreak", () => {
    const { xml } = stringifyParagraphProperties({
      thematicBreak: true,
      border: { bottom: { color: "FF0000", size: 8, space: 1, style: "double" } },
    });
    expect(xml?.match(/<w:pBdr/g)).toHaveLength(1);
    expect(xml).toContain('w:color="FF0000"');
    expect(xml).toContain('w:val="double"');
  });
});
