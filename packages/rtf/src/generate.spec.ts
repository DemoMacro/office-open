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

  it("round-trips left, right, and first headers and footers", () => {
    const options: DocumentOptions = {
      sections: [
        {
          headers: {
            default: [{ paragraph: { children: [{ text: "Right" }] } }],
            even: [{ paragraph: { children: [{ text: "Left" }] } }],
            first: [{ paragraph: { children: [{ text: "First" }] } }],
          },
          footers: {
            default: [{ paragraph: { children: [{ text: "Bottom" }] } }],
            even: [{ paragraph: { children: [{ text: "Even" }] } }],
          },
          children: [],
        },
      ],
    };
    expect(roundTrip(decoder.decode(generateDocument(options)))).toEqual(options);
  });

  it("round-trips footnote references at their source positions", () => {
    const source = String.raw`{\rtf1Before{\footnote Note}After\par}`;
    const options = parseDocument(source);
    expect(options.sections[0]?.children?.[0]).toEqual({
      paragraph: { children: [{ text: "Before" }, { footnoteReference: 0 }, { text: "After" }] },
    });
    expect(roundTrip(decoder.decode(generateDocument(options)))).toEqual(options);
  });

  it("projects styles, fonts, colors, and character formatting", () => {
    const source = String.raw`{\rtf1{\stylesheet{\s0 Heading;}{\cs0 Term;\strike\super\lang1033\caps\scaps\v\charscalex90}}{\fonttbl{\f0 Arial;}}{\colortbl;\red255\green0\blue0;}\s0\f0\fs20\cf1\b\i\ul Heading\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects run language controls", () => {
    const source = String.raw`{\rtf1{\lang1033 Word}{\lang1036 Autre}\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects hidden text controls", () => {
    const source = String.raw`{\rtf1{\v hidden}visible\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects numbering, paragraphs, tabs, breaks, bookmarks, and inline children", () => {
    const source = String.raw`{\rtf1\qc\ls1\ilvl0{\pn\pndec}First{\*\bkmkstart Mark}A{\*\bkmkend Mark}\tab\line\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("round-trips complete numbering tables and pntext shadow controls", () => {
    const source = String.raw`{\rtf1{\listtable{\list\listtemplateid9\listhybrid{\listlevel\levelnfc2\leveljc1\levelstartat3\levelindent720{\leveltext\'02\'00.;}{\levelnumbers;}}{\listname rtf-list-template-9;}\listid77}}{\listoverridetable{\listoverride\listid77\listoverridecount0\ls4}}\ls4\ilvl0\par Item}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects flat and nested tables", () => {
    const source = String.raw`{\rtf1\trowd\cellx1000 Outer{\trowd\cellx500 Inner\cell}\nestrow End\cell\row}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("round-trips table row and cell properties", () => {
    const source = String.raw`{\rtf1{\colortbl;\red255\green0\blue0;}\trowd\trhdr\trkeep\trh300\trqc\clcbpat1\clvertalc\clvmgf\cellx1000\clvmrg\cellx2000 First\cell\cell\row}`;
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

  it("projects wrapped floating shape anchors", () => {
    const source = String.raw`{\rtf1{\shp{\*\shpinst\shpbxpage\shpbypage\shpleft100\shptop100\shpwidth1000\shpheight500\shpfblwtxt1\shplockanchor\shpz7\shpwr2\shpwrk3\shptxt Shape}}\par}`;
    expect(roundTrip(source)).toEqual(parseDocument(source));
  });

  it("projects revision IDs and hexadecimal object data", () => {
    const source = String.raw`{\rtf1\insrsid27113541 Revised{\object{\objclass Test.Class}{\objdata 48656C6C6F}}\par}`;
    const parsed = parseDocument(source);
    expect(parsed.sections?.[0]?.children[0]).toMatchObject({
      paragraph: {
        children: [
          { additionRsid: "019DB845" },
          {
            additionRsid: "019DB845",
            children: [
              {
                object: {
                  embed: { progId: "Test.Class", data: new Uint8Array([72, 101, 108, 108, 111]) },
                },
              },
            ],
          },
        ],
      },
    });
    expect(roundTrip(source)).toEqual(parsed);
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
