import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { PptxWriteContext } from "../../context";
import { slideMasterDesc } from "./slide-master";
import {
  DEFAULT_TEXT_LIST_STYLE,
  parseTextListStyle,
  stringifyTextListStyle,
} from "./text-list-style";

describe("text-list-style master emit contract", () => {
  it("fresh master txStyles equals the structured default stringify", () => {
    // slideMasterDesc must emit stringifyTextListStyle(DEFAULT_TEXT_LIST_STYLE)
    // — no divergent hardcoded constant. Byte-stability of the fresh master
    // output (and thus the 16-round-trip fidelity gate) depends on this.
    const writeCtx = new PptxWriteContext();
    writeCtx.slideWidth = 12192000;
    const master = slideMasterDesc.stringify({}, writeCtx)!;
    const start = master.indexOf("<p:txStyles>");
    const end = master.indexOf("</p:txStyles>", start) + "</p:txStyles>".length;
    const emittedBlock = master.slice(start, end);

    const structuredBlock = `<p:txStyles>${stringifyTextListStyle(DEFAULT_TEXT_LIST_STYLE)}</p:txStyles>`;

    expect(structuredBlock.length).toBe(emittedBlock.length);
    expect(structuredBlock).toBe(emittedBlock);
  });
});

describe("text-list-style round-trip", () => {
  // nativeTypeAttributes:true mirrors the real OPC parse path, which coerces
  // "1"/"0" to numbers. Boolean-attribute parsing must survive that coercion
  // (a strict `=== "1"` check silently fails on numeric 1 and breaks the
  // 16-round-trip byte-fidelity gate).
  const parseCoerced = (xml: string) => {
    const el = parseXml(xml, { nativeTypeAttributes: true }).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    return el;
  };

  it("parse(stringify(default)) is stable under attribute coercion", () => {
    const xml = `<p:txStyles>${stringifyTextListStyle(DEFAULT_TEXT_LIST_STYLE)}</p:txStyles>`;
    const parsed = parseTextListStyle(parseCoerced(xml));
    const restringified = stringifyTextListStyle(parsed);
    expect(restringified).toBe(stringifyTextListStyle(DEFAULT_TEXT_LIST_STYLE));
  });

  it("preserves title/body/other levels and run properties", () => {
    const xml = `<p:txStyles>${stringifyTextListStyle(DEFAULT_TEXT_LIST_STYLE)}</p:txStyles>`;
    const parsed = parseTextListStyle(parseCoerced(xml));

    expect(parsed.title?.levels?.[0]?.defaultRun?.size).toBe(44);
    expect(parsed.title?.levels?.[0]?.bullet?.type).toBe("none");
    expect(parsed.body?.levels?.[1]?.defaultRun?.size).toBe(24);
    expect(parsed.body?.levels?.[1]?.bullet?.type).toBe("char");
    expect(parsed.body?.levels?.[1]?.spaceBeforePoints).toBe(5);
    expect(parsed.other?.emptyDefaultParagraph).toBe(true);
    expect(parsed.other?.levels?.[8]?.indent).toBe(-228600);
    expect(parsed.other?.levels?.[0]?.defaultRun?.latin).toBe("+mn-lt");
    // Boolean attributes that nativeTypeAttributes coerces "1"→1.
    expect(parsed.title?.levels?.[0]?.eastAsianLineBreak).toBe(true);
    expect(parsed.title?.levels?.[0]?.hangingPunctuation).toBe(true);
    expect(parsed.title?.levels?.[0]?.latinLineBreak).toBe(false);
  });
});
