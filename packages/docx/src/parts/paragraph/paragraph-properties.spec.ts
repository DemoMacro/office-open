import { describe, expect, it } from "vite-plus/test";

import { stringifyParagraphProperties } from "./stringify";

describe("stringifyParagraphProperties single-writer channels", () => {
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
