import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vite-plus/test";

import { parseDocument } from ".";
import { RtfGenerateError } from "./errors";
import { generateDocument } from "./generate";

const decoder = new TextDecoder();

function roundTrip(source: string): DocumentOptions {
  const options = parseDocument(source);
  return parseDocument(decoder.decode(generateDocument(options)));
}

describe("generateDocument RTF projections", () => {
  it("projects document text, metadata, sections, headers, footers, and footnotes", () => {
    const source = String.raw`{\rtf1{\info{\title Title}{\author Author}}{\header Header}{\footer Footer}{\footnote Note}\par Body\sect Second\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects styles, fonts, colors, and character formatting", () => {
    const source = String.raw`{\rtf1{\stylesheet{\s0 Heading;}{\cs0 Term;}}{\fonttbl{\f0 Arial;}}{\colortbl;\red255\green0\blue0;}\s0\f0\fs20\cf1\b\i\ul Heading\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects numbering, paragraphs, tabs, breaks, bookmarks, and inline children", () => {
    const source = String.raw`{\rtf1\qc\ls1\ilvl0{\pn\pndec}First{\*\bkmkstart Mark}A{\*\bkmkend Mark}\tab\line\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects flat and nested tables", () => {
    const source = String.raw`{\rtf1\trowd\cellx1000\cellx2000 Outer{\trowd\cellx500 Inner\cell}\nestrow End\cell\row}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects fields and hyperlinks", () => {
    const source = String.raw`{\rtf1{\field{\*\fldinst HYPERLINK "https://example.com"}{\fldrslt Link}}{\field{\*\fldinst PAGE}{\fldrslt 1}}\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects pictures", () => {
    const source = String.raw`{\rtf1{\pict\pngblip\picw10\pich5 89504e470d0a1a0a}}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects shapes", () => {
    const source = String.raw`{\rtf1{\shpinst\shpleft100\shptop100\shpwidth1000\shpheight500\shptxt Shape}}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("rejects unsupported document fields with structured context", () => {
    const options: DocumentOptions = {
      sections: [{ children: [{ paragraph: { children: [{ text: "Body" }] } }] }],
      comments: [],
    };
    expect(() => generateDocument(options)).toThrow(RtfGenerateError);
    expect(() => generateDocument(options)).toThrow(
      expect.objectContaining({
        context: expect.objectContaining({ part: "document", path: "options", name: "comments" }),
      }),
    );
  });
});
