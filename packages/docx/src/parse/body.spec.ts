import { parse as parseXml } from "@office-open/xml";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../context";
import { generateDocumentSync } from "../generate";
import { parseDocumentSync } from "../parse";
import { parseSectionChild } from "./body";

const readCtx = {} as unknown as DocxReadContext;

/** The shape office-open's own textbox stringifier emits: w:p > w:r > w:pict. */
const TEXTBOX_PICT =
  '<w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202" style="width:120pt;height:24pt">' +
  "<v:textbox><w:txbxContent><w:p><w:r><w:t>In box</w:t></w:r></w:p></w:txbxContent></v:textbox>" +
  "</v:shape></w:pict>";

function parseFirstChild(xml: string) {
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseSectionChild(el, readCtx);
}

describe("parseSectionChild run-wrapped textboxes", () => {
  it("promotes the run-wrapped pict the textbox stringifier emits", () => {
    const child = parseFirstChild(
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${TEXTBOX_PICT}</w:r></w:p>`,
    );
    expect("textbox" in (child as object)).toBe(true);
    expect(JSON.stringify(child)).toContain("In box");
  });

  it("keeps a bare pict direct child working", () => {
    const child = parseFirstChild(`<w:p>${TEXTBOX_PICT}</w:p>`);
    expect("textbox" in (child as object)).toBe(true);
  });

  it("does not promote a run that mixes text with a pict", () => {
    const child = parseFirstChild(`<w:p><w:r><w:t>Before</w:t>${TEXTBOX_PICT}</w:r></w:p>`);
    expect("textbox" in (child as object)).toBe(false);
    expect(JSON.stringify(child)).toContain("Before");
  });

  it("does not promote a paragraph that carries extra runs", () => {
    const child = parseFirstChild(
      `<w:p><w:r><w:t>Caption</w:t></w:r><w:r>${TEXTBOX_PICT}</w:r></w:p>`,
    );
    expect("textbox" in (child as object)).toBe(false);
  });
});

describe("textbox generate → parse round trip", () => {
  it("keeps a generated textbox as a textbox on re-parse", () => {
    const first = generateDocumentSync({
      sections: [
        {
          children: [
            { textbox: { children: [{ paragraph: { children: [{ text: "In box" }] } }] } },
          ],
        },
      ],
    });
    const opts = parseDocumentSync(first);
    const children = opts.sections?.[0]?.children ?? [];
    expect(children.some((c) => "textbox" in c)).toBe(true);

    const second = generateDocumentSync(opts);
    const xml = new TextDecoder().decode(unzipSync(second)["word/document.xml"]!);
    expect(xml).toContain("v:textbox");
    expect(xml).toContain("In box");
  });
});
