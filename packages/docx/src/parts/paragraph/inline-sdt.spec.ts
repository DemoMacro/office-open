import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

// Inline SDTs never touch the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

function parseParagraphXml(inner: string): { children?: unknown[] } {
  const doc = parseXml(`<w:p ${W_NS}>${inner}</w:p>`);
  return parseParagraph(doc.elements![0], readCtx) as { children?: unknown[] };
}

function findInlineSdt(opts: { children?: unknown[] }): Record<string, unknown> | undefined {
  return opts.children?.find(
    (c) => c !== null && typeof c === "object" && "sdt" in (c as Record<string, unknown>),
  ) as Record<string, unknown> | undefined;
}

describe("inline SDT (CT_SdtRun) parse", () => {
  it("parses a plain-text inline SDT with run content", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:tag w:val="name"/><w:text/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>Jane</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect(child).toBeDefined();
    const sdt = child!.sdt as Record<string, unknown>;
    expect(sdt.properties).toMatchObject({ alias: "Name", tag: "name" });
    expect(sdt.children).toBeDefined();
    expect((sdt.children as unknown[]).length).toBeGreaterThan(0);
  });

  it("parses an inline SDT alongside surrounding runs (position preserved)", () => {
    const opts = parseParagraphXml(
      "<w:r><w:t>before </w:t></w:r>" +
        '<w:sdt><w:sdtPr><w:alias w:val="X"/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>mid</w:t></w:r></w:sdtContent></w:sdt>" +
        "<w:r><w:t> after</w:t></w:r>",
    );
    const child = findInlineSdt(opts);
    expect(child).toBeDefined();
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({ alias: "X" });
  });

  it("parses inline SDT end properties (w:sdtEndPr)", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="E"/></w:sdtPr>' +
        "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>" +
        "<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>",
    );
    const child = findInlineSdt(opts);
    expect((child!.sdt as Record<string, unknown>).endProperties).toBeDefined();
  });

  it("parses an inline checkbox SDT (w14:checkbox) — sdtContent is a symbol run, not a w:p", () => {
    const opts = parseParagraphXml(
      `<w:sdt><w:sdtPr ${W14}><w:alias w:val="C"/><w14:checkbox>` +
        '<w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>' +
        '<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>',
    );
    const child = findInlineSdt(opts);
    expect(child).toBeDefined();
    expect((child!.sdt as Record<string, unknown>).properties).toMatchObject({
      alias: "C",
      checkbox: { checked: true },
    });
  });

  it("parses a nested inline SDT (SDT inside SDT content)", () => {
    const opts = parseParagraphXml(
      '<w:sdt><w:sdtPr><w:alias w:val="Outer"/></w:sdtPr>' +
        '<w:sdtContent><w:sdt><w:sdtPr><w:alias w:val="Inner"/></w:sdtPr>' +
        "<w:sdtContent><w:r><w:t>deep</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt>",
    );
    const outer = findInlineSdt(opts);
    const outerChildren = (outer!.sdt as Record<string, unknown>).children as unknown[];
    const inner = outerChildren.find(
      (c) => c !== null && typeof c === "object" && "sdt" in (c as Record<string, unknown>),
    ) as Record<string, unknown> | undefined;
    expect(inner).toBeDefined();
    expect((inner!.sdt as Record<string, unknown>).properties).toMatchObject({ alias: "Inner" });
  });
});
