import { describe, expect, it } from "vitest";

import { stringify } from "../src/stringify";
import type { Element } from "../src/types";

describe("stringify", () => {
  it("stringifies a simple element", () => {
    const el: Element = {
      elements: [{ type: "element", name: "w:t", elements: [{ type: "text", text: "Hello" }] }],
    };
    expect(stringify(el)).toBe("<w:t>Hello</w:t>");
  });

  it("stringifies an element with attributes", () => {
    const el: Element = {
      elements: [{ type: "element", name: "w:pStyle", attributes: { "w:val": "Title" } }],
    };
    expect(stringify(el)).toBe('<w:pStyle w:val="Title"/>');
  });

  it("stringifies nested elements", () => {
    const el: Element = {
      elements: [
        {
          type: "element",
          name: "w:p",
          attributes: { "w:val": "1" },
          elements: [
            {
              type: "element",
              name: "w:r",
              elements: [
                {
                  type: "element",
                  name: "w:t",
                  elements: [{ type: "text", text: "Hello" }],
                },
              ],
            },
          ],
        },
      ],
    };
    expect(stringify(el)).toBe('<w:p w:val="1"><w:r><w:t>Hello</w:t></w:r></w:p>');
  });

  it("writes XML declaration", () => {
    const el: Element = {
      declaration: { attributes: { version: "1.0", encoding: "UTF-8", standalone: "yes" } },
      elements: [{ type: "element", name: "root" }],
    };
    expect(stringify(el)).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>');
  });

  it("preserves supported XML declaration versions", () => {
    const el: Element = {
      declaration: { attributes: { version: "1.1" } },
      elements: [{ type: "element", name: "root" }],
    };
    expect(stringify(el)).toBe('<?xml version="1.1"?><root/>');
    expect(() =>
      stringify({
        declaration: { attributes: { version: "2.0" } },
        elements: [{ type: "element", name: "root" }],
      }),
    ).toThrowError("Unsupported XML declaration version 2.0.");
  });

  it("handles empty elements (self-closing)", () => {
    const el: Element = { elements: [{ type: "element", name: "root" }] };
    expect(stringify(el)).toBe("<root/>");
  });

  it("forces closing tag for xml:space=preserve", () => {
    const el: Element = {
      elements: [
        {
          type: "element",
          name: "w:t",
          attributes: { "xml:space": "preserve" },
        },
      ],
    };
    expect(stringify(el)).toBe('<w:t xml:space="preserve"></w:t>');
  });

  it("handles CDATA", () => {
    const el: Element = {
      elements: [{ type: "element", name: "root", elements: [{ type: "cdata", cdata: "data" }] }],
    };
    expect(stringify(el)).toBe("<root><![CDATA[data]]></root>");
  });

  it("handles comments", () => {
    const el: Element = {
      elements: [
        { type: "element", name: "root", elements: [{ type: "comment", comment: "note" }] },
      ],
    };
    expect(stringify(el)).toBe("<root><!--note--></root>");
  });

  it("writes and ignores processing instructions", () => {
    const el: Element = {
      elements: [
        { type: "instruction", instruction: 'mso-application progid="Word.Document"' },
        { type: "element", name: "root" },
      ],
    };
    expect(stringify(el)).toBe('<?mso-application progid="Word.Document"?><root/>');
    expect(stringify(el, { ignoreInstruction: true })).toBe("<root/>");
  });

  it("preserves valid raw deferred content", () => {
    const el: Element = {
      elements: [{ type: "element", name: "root", raw: '<child value="1" /><child/>' }],
    };
    expect(stringify(el)).toBe('<root><child value="1" /><child/></root>');
  });

  it("rejects malformed raw deferred content", () => {
    const cases: Array<[string, Element]> = [
      ["unclosed element", { elements: [{ type: "element", name: "root", raw: "<child>" }] }],
      [
        "mismatched element",
        { elements: [{ type: "element", name: "root", raw: "<child></other>" }] },
      ],
      [
        "unclosed attribute",
        { elements: [{ type: "element", name: "root", raw: '<child value="1>' }] },
      ],
    ];
    for (const [, el] of cases) {
      expect(() => stringify(el)).toThrowError("Malformed XML: raw content for <root>");
    }
  });

  it("rejects reserved XML instruction and comment content", () => {
    expect(() =>
      stringify({
        elements: [{ type: "instruction", instruction: "a?>b" }],
      }),
    ).toThrowError("Malformed XML: processing instruction contains its closing delimiter.");
    expect(() => stringify({ elements: [{ type: "comment", comment: "a--b" }] })).toThrowError(
      "Malformed XML: comment contains the reserved double hyphen.",
    );
    expect(() => stringify({ elements: [{ type: "comment", comment: "a-" }] })).toThrowError(
      "Malformed XML: comment ends with a hyphen.",
    );
  });

  it("omits attributes when ignoreAttributes is enabled", () => {
    const el: Element = {
      elements: [{ type: "element", name: "root", attributes: { flag: 1 } }],
    };
    expect(stringify(el, { ignoreAttributes: true })).toBe("<root/>");
  });

  it("handles text nodes with special characters", () => {
    const el: Element = {
      elements: [{ type: "element", name: "w:t", elements: [{ type: "text", text: "a & b" }] }],
    };
    expect(stringify(el)).toBe("<w:t>a &amp; b</w:t>");
  });

  it("escapes ampersand in text content", () => {
    const el: Element = {
      elements: [{ type: "element", name: "w:t", elements: [{ type: "text", text: "&amp;" }] }],
    };
    expect(stringify(el)).toBe("<w:t>&amp;amp;</w:t>");
  });

  it("applies attributeValueFn", () => {
    const el: Element = {
      elements: [
        {
          type: "element",
          name: "w:p",
          attributes: { "w:val": "a&b" },
        },
      ],
    };
    const result = stringify(el, {
      attributeValueFn: (value) => value.replace(/&/g, "&amp;"),
    });
    expect(result).toBe('<w:p w:val="a&amp;b"/>');
  });

  it("converts numeric attribute values to string", () => {
    const el: Element = {
      elements: [{ type: "element", name: "w:val", attributes: { count: 42 } }],
    };
    expect(stringify(el)).toBe('<w:val count="42"/>');
  });

  it("handles element with no name", () => {
    const el: Element = { elements: [{ type: "text", text: "just text" }] };
    expect(stringify(el)).toBe("just text");
  });
});
