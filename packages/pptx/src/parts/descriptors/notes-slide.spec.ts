import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { PptxWriteContext } from "../../context";
import { notesSlideDesc } from "./notes-slide";
import type { NotesSlideOptions } from "./notes-slide";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as PptxWriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

/** stringify → parse → stringify must be byte-stable (round-trip fidelity). */
function restringify(opts: NotesSlideOptions): string {
  const original = notesSlideDesc.stringify(opts, writeCtx)!;
  const el = parseXml(original).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  const parsed = notesSlideDesc.parse(el, readCtx);
  return notesSlideDesc.stringify(parsed, writeCtx)!;
}

describe("notesSlideDesc round-trip", () => {
  it("simple text builds a fresh notes slide (sldImg + body placeholders)", () => {
    const xml = notesSlideDesc.stringify({ text: "Speaker notes content" }, writeCtx)!;
    expect(xml).toContain('type="sldImg"');
    expect(xml).toContain('type="body"');
    expect(xml).toContain("Speaker notes content");
  });

  it("round-trips notes text byte-stable", () => {
    const restringified = restringify({ text: "Speaker notes content" });
    expect(restringified).toContain("Speaker notes content");
    // sldImg placeholder has no txBody; it must stay txBody-less on re-emit.
    const sldImgStart = restringified.indexOf('type="sldImg"');
    const sldImgSpEnd = restringified.indexOf("</p:sp>", sldImgStart);
    expect(restringified.slice(sldImgStart, sldImgSpEnd)).not.toContain("<p:txBody");
  });

  it("round-trips empty notes byte-stable", () => {
    const original = notesSlideDesc.stringify({}, writeCtx)!;
    const el = parseXml(original).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const parsed = notesSlideDesc.parse(el, readCtx);
    expect(notesSlideDesc.stringify(parsed, writeCtx)!).toBe(original);
  });

  it("round-trips notes with special characters byte-stable", () => {
    const original = notesSlideDesc.stringify({ text: '<b>Bold & "quoted"' }, writeCtx)!;
    expect(original).toContain("&lt;b&gt;Bold &amp; &quot;quoted&quot;");
    const el = parseXml(original).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const parsed = notesSlideDesc.parse(el, readCtx);
    expect(notesSlideDesc.stringify(parsed, writeCtx)!).toBe(original);
  });
});
