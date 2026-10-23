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

  it("projects valid W3CDTF and Date metadata timestamps", () => {
    const generated = decoder.decode(
      generateDocument({
        sections: [{ children: [] }],
        created: "2026-10-19T08:09:10Z",
        createdType: "dcterms:W3CDTF",
        modified: new Date(Date.UTC(2026, 8, 1, 7, 6, 5)) as unknown as string,
        modifiedType: false,
        lastPrinted: "2026-10-19T02:03:04+08:00",
      }),
    );
    expect(generated).toContain(String.raw`{\creatim\yr2026\mo10\dy19\hr08\min09\sec10}`);
    expect(generated).toContain(String.raw`{\revtim\yr2026\mo09\dy01\hr07\min06\sec05}`);
    expect(generated).toContain(String.raw`{\printim\yr2026\mo10\dy18\hr18\min03\sec04}`);
  });

  it("rejects invalid and lossy metadata timestamps", () => {
    const sections = [{ children: [] }] as DocumentOptions["sections"];
    expect(() =>
      generateDocument({ sections, created: "19 October 2026" as unknown as string }),
    ).toThrow(RtfGenerateError);
    expect(() =>
      generateDocument({
        sections,
        modified: new Date(Date.UTC(2026, 0, 1, 0, 0, 1, 500)) as unknown as string,
      }),
    ).toThrow(RtfGenerateError);
  });

  it("projects lastModifiedBy and endnote references", () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [{ paragraph: { children: [{ text: "Body" }, { endnoteReference: 0 }] } }],
        },
      ],
      lastModifiedBy: "Editor",
      endnotes: [
        {
          id: 0,
          children: [{ paragraph: { children: [{ text: "End" }] } }],
        },
      ],
    };
    const generated = decoder.decode(generateDocument(options));
    expect(generated).toContain(String.raw`{\operator Editor}`);
    expect(generated).toContain(String.raw`{\footnote\ftnalt End\par`);
    expect(roundTrip(generated)).toEqual(options);
  });

  it("projects document language and generating application", () => {
    const options: DocumentOptions = {
      sections: [{ children: [] }],
      language: "en-US",
      application: "office-open",
    };
    const generated = decoder.decode(generateDocument(options));
    expect(generated).toContain(String.raw`{\generator office-open}`);
    expect(generated).toContain(String.raw`\deflang1033 `);
    expect(roundTrip(generated)).toEqual(options);
  });

  it("rejects unrepresentable OOXML package state", () => {
    const sections = [{ children: [{ paragraph: { children: [{ text: "Body" }] } }] }];
    expect(() =>
      generateDocument({
        sections,
        rawParts: [
          {
            path: "custom/raw.bin",
            data: new Uint8Array([1, 2, 3]),
            contentType: "application/octet-stream",
          },
        ],
      }),
    ).toThrow(RtfGenerateError);
    expect(() =>
      generateDocument({
        sections,
        passthroughRelationships: [
          {
            source: "word/document.xml",
            relationshipType: "http://example.com/raw",
            target: "custom/raw.bin",
            rId: "rId8",
          },
        ],
      }),
    ).toThrow(RtfGenerateError);
  });

  it("projects expressible document settings and rejects unsupported settings", () => {
    const options: DocumentOptions = {
      sections: [{ children: [] }],
      settings: {
        defaultTabStop: "36pt",
        view: "print",
        zoom: { percent: "150%" },
        trackRevisions: true,
        doNotTrackMoves: false,
      },
    };
    const generated = decoder.decode(generateDocument(options));
    expect(generated).toContain(String.raw`\deftab720 `);
    expect(generated).toContain(String.raw`\viewkind1 `);
    expect(generated).toContain(String.raw`\viewscale150 `);
    expect(generated).toContain(String.raw`\revisions `);
    expect(generated).toContain(String.raw`\trackmoves `);
    expect(() =>
      generateDocument({
        sections: [{ children: [] }],
        settings: { updateFields: true },
      }),
    ).toThrow(
      expect.objectContaining({
        context: expect.objectContaining({
          part: "settings",
          path: "settings",
          name: "updateFields",
        }),
      }),
    );
  });

  it("projects section-child bookmark starts", () => {
    const generated = decoder.decode(
      generateDocument({
        sections: [
          {
            children: [
              { bookmarkStart: { id: 7, name: "Section" } },
              { paragraph: { children: [{ text: "Body" }] } },
            ],
          },
        ],
      }),
    );
    expect(generated.indexOf(String.raw`{\*\bkmkstart Section}`)).toBeLessThan(
      generated.indexOf("Body"),
    );
  });

  it("projects paragraph indents", () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                indent: { left: 720, right: 240, firstLine: 120 },
                children: [{ text: "Body" }],
              },
            },
          ],
        },
      ],
    };
    const generated = decoder.decode(generateDocument(options));
    expect(generated).toContain(String.raw`\li720 \ri240 \fi120`);
    expect(parseDocument(generated)).toEqual(options);
  });

  it("projects formatted runs containing hyperlinks", () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    bold: true,
                    children: [{ hyperlink: { url: "https://example.com", children: ["Link"] } }],
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const generated = decoder.decode(generateDocument(options));
    expect(generated).toContain(String.raw`HYPERLINK "https://example.com"`);
    expect(parseDocument(generated)).toEqual(options);
  });

  it("maps declared style ids to stable RTF style codes", () => {
    const generated = decoder.decode(
      generateDocument({
        sections: [
          {
            children: [
              {
                paragraph: {
                  style: "Heading",
                  children: [{ style: "Emphasis", children: [{ text: "Body" }] }],
                },
              },
            ],
          },
        ],
        styles: {
          paragraphStyles: [{ id: "Heading", name: "Heading" }],
          characterStyles: [{ id: "Emphasis", name: "Emphasis" }],
        },
      }),
    );
    expect(generated).toContain(String.raw`{\stylesheet\s0`);
    expect(generated).toContain(String.raw`\cs1`);
    expect(generated).toContain(String.raw`\s0 {\cs1 Body}`);
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

  it("projects paragraph text shorthand with preserved spaces", () => {
    const generated = decoder.decode(
      generateDocument({
        sections: [{ children: [{ paragraph: { text: " Body ", preserveSpace: true } }] }],
      }),
    );
    expect(parseDocument(generated).sections[0]?.children?.[0]).toEqual({
      paragraph: { children: [{ text: " Body " }] },
    });
  });
  it("rejects paragraph text and children together", () => {
    const options: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                text: "Body",
                children: [{ text: "Children" }],
              },
            },
          ],
        },
      ],
    };
    expect(() => generateDocument(options)).toThrow(RtfGenerateError);
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

  it("round-trips font family and character-set metadata", () => {
    const source = String.raw`{\rtf1{\fonttbl{\f0\fswiss\fcharset0 Arial;}}\f0 Text\par}`;
    const parsed = parseDocument(source);
    expect(parsed.fonts).toEqual([{ name: "Arial", family: "swiss", characterSet: "00" }]);
    const generated = decoder.decode(generateDocument(parsed));
    expect(generated).toContain(String.raw`{\fonttbl{\f0\fswiss\fcharset0 Arial;}}`);
    expect(parseDocument(generated)).toEqual(parsed);
  });

  it("rejects unknown RTF font families structurally", () => {
    expect(() =>
      generateDocument({
        fonts: [{ name: "Custom", family: "unknown" }],
        sections: [],
      }),
    ).toThrow(RtfGenerateError);
  });

  it("rejects duplicate RTF font-table definitions", () => {
    expect(() =>
      generateDocument({
        fonts: [{ name: "Arial" }, { name: "Arial", family: "swiss" }],
        sections: [],
      }),
    ).toThrow(RtfGenerateError);
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

  it("encodes astral-plane text as complete Unicode code points", () => {
    const generated = decoder.decode(
      generateDocument({
        sections: [{ children: [{ paragraph: { children: [{ text: "𝄞" }] } }] }],
      }),
    );
    expect(generated).toContain("\\u-10188?\\u-8930?");
    expect(parseDocument(generated).sections?.[0]?.children?.[0]).toEqual({
      paragraph: { children: [{ text: "𝄞" }] },
    });
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
