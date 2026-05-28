// Round-trip demo — one document with all content types, export → parse → verify → re-export.

import * as fs from "fs";

import { Document, Packer, parseDocument } from "@office-open/docx";
import type { SectionOptions } from "@office-open/docx";

// ─── Helpers ───────────────────────────────────────────────────────────────

let pass = 0;
let fail = 0;

function assert(label: string, condition: boolean) {
  if (condition) {
    pass++;
    console.log(`  PASS: ${label}`);
  } else {
    fail++;
    console.log(`  FAIL: ${label}`);
  }
}

// ─── Main ──────────────────────────────────────────────────────────────────

async function main() {
  // 1. Build a single document with every content type
  const sections: SectionOptions[] = [
    // Section: paragraphs, tables, nested tables, SDT, headings, TOC, headers/footers
    {
      headers: {
        default: [{ paragraph: { children: [{ text: "Header", bold: true }] } }],
      },
      footers: {
        default: [{ paragraph: { children: ["Footer text"] } }],
      },
      children: [
        // String paragraph
        { paragraph: "I am a string paragraph" },

        // Paragraph with mixed run children
        {
          paragraph: {
            children: [
              "String child",
              { text: " + bold", bold: true },
              { text: " + italic", italics: true },
            ],
          },
        },

        // Centered paragraph
        {
          paragraph: {
            alignment: "center",
            children: [{ text: "Centered text", bold: true, size: 32 }],
          },
        },

        // Comment
        {
          paragraph: {
            children: [
              "This paragraph has a ",
              { commentRangeStart: 0 },
              "comment",
              { commentRangeEnd: 0 },
              { commentReference: 0 },
              " attached.",
            ],
          },
        },

        // Hyperlink
        {
          paragraph: {
            children: [
              "Visit ",
              {
                text: "our website",
                hyperlink: { link: "https://example.com", tooltip: "Example" },
              },
              " for more.",
            ],
          },
        },

        // Bookmark
        {
          paragraph: {
            children: [
              { bookmarkStart: { id: 42, name: "my-bookmark" } },
              "This is a bookmarked paragraph.",
              { bookmarkEnd: 42 },
            ],
          },
        },

        // Table of Contents
        { toc: { hyperlink: true, headingStyleRange: "1-3" } },

        // Headings
        { paragraph: { heading: "Heading1", children: ["First Heading"] } },
        { paragraph: { children: ["Content under first heading."] } },
        { paragraph: { heading: "Heading2", children: ["Second Heading"] } },
        { paragraph: { children: ["Content under second heading."] } },

        // Table
        {
          table: {
            rows: [
              {
                children: [
                  { children: [{ paragraph: "A1" }] },
                  { children: [{ paragraph: "B1" }] },
                  { children: [{ paragraph: "C1" }] },
                ],
              },
              {
                children: [
                  {
                    children: [
                      { paragraph: { children: [{ text: "Underlined", underline: {} }] } },
                    ],
                  },
                  {
                    children: [
                      { paragraph: "Cell with" },
                      {
                        paragraph: { children: ["multiple", { text: " paragraphs", bold: true }] },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        },

        // Nested table
        {
          table: {
            rows: [
              {
                children: [
                  {
                    children: [
                      { paragraph: "Outer cell" },
                      {
                        table: {
                          rows: [
                            {
                              children: [
                                { children: [{ paragraph: "Inner A" }] },
                                { children: [{ paragraph: "Inner B" }] },
                              ],
                            },
                          ],
                        },
                      },
                    ],
                  },
                  { children: [{ paragraph: "Sibling" }] },
                ],
              },
            ],
          },
        },

        // SDT block
        {
          sdt: {
            properties: { richText: true, alias: "BlockContent", tag: "block-content" },
            children: [
              { paragraph: { children: ["Block-level content control."] } },
              { paragraph: { children: ["Multiple paragraphs inside SDT."] } },
            ],
          },
        },

        // ComboBox SDT
        {
          sdt: {
            properties: {
              alias: "Color",
              tag: "combo-color",
              comboBox: {
                items: [
                  { displayText: "Red", value: "red" },
                  { displayText: "Blue", value: "blue" },
                ],
                lastValue: "Red",
              },
            },
            children: [{ paragraph: { children: ["Red"] } }],
          },
        },

        // Textbox
        {
          textbox: {
            style: { width: "4in", height: "1in" },
            children: [
              { paragraph: { children: [{ text: "Textbox paragraph", bold: true }] } },
              { paragraph: { children: ["Second line in textbox"] } },
            ],
          },
        },
      ],
    },

    // Section: chart
    {
      children: [
        { paragraph: { children: [{ text: "Chart (JSON API)", bold: true, size: 28 }] } },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "column",
                  data: {
                    categories: ["Q1", "Q2", "Q3", "Q4"],
                    series: [
                      { name: "2024", values: [120, 150, 180, 200] },
                      { name: "2025", values: [140, 170, 210, 250] },
                    ],
                  },
                  title: "Quarterly Revenue",
                  transformation: { width: 500, height: 300 },
                },
              },
            ],
          },
        },
      ],
    },

    // Section: smartArt
    {
      children: [
        { paragraph: { children: [{ text: "SmartArt (JSON API)", bold: true, size: 28 }] } },
        {
          paragraph: {
            children: [
              {
                smartArt: {
                  data: {
                    nodes: [
                      { text: "Plan", children: [{ text: "Define scope" }] },
                      { text: "Build" },
                      { text: "Deploy" },
                    ],
                  },
                  transformation: { width: 450, height: 250 },
                  layout: "process1",
                  style: "simple1",
                  color: "accent1_2",
                },
              },
            ],
          },
        },
      ],
    },

    // Section: altChunk
    {
      children: [
        { paragraph: { children: [{ text: "AltChunk (JSON API)", bold: true, size: 28 }] } },
        {
          altChunk: {
            data: "<html><body><p>This is <b>embedded HTML</b> via JSON API.</p></body></html>",
            contentType: "text/html",
            extension: "html",
          },
        },
      ],
    },

    // Section: image
    {
      children: [
        { paragraph: { children: [{ text: "Image (JSON API)", bold: true, size: 28 }] } },
        {
          paragraph: {
            children: [
              {
                image: {
                  type: "png",
                  data: fs.readFileSync("./demo/images/dog.png"),
                  transformation: { width: 200, height: 200 },
                },
              },
            ],
          },
        },
      ],
    },

    // Section: math (JSON API with complex structures), symbol, breaks
    {
      children: [
        {
          paragraph: {
            children: [
              "The formula E=mc",
              { math: { children: [{ text: "2" }] } },
              " uses a ",
              { symbolRun: { char: "F021", symbolfont: "Wingdings" } },
              " symbol.",
            ],
          },
        },
        // Complex math: fraction
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    "x",
                    " = ",
                    {
                      fraction: {
                        numerator: ["a + b"],
                        denominator: ["c"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        // Complex math: sum with superScript
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    {
                      sum: {
                        children: [{ superScript: { children: ["x"], superScript: ["2"] } }],
                        subScript: ["i=0"],
                        superScript: ["n"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: ["Before page break", { pageBreak: true }, "After page break"],
          },
        },
        {
          paragraph: {
            children: ["Column 1", { columnBreak: true }, "Column 2"],
          },
        },
        {
          paragraph: {
            children: ["With footnote", { footnoteReference: 1 }],
          },
        },
      ],
    },

    // Section: rich formatting for value comparison
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Styled", bold: true, italics: true, size: 28, color: "FF0000" }],
          },
        },
      ],
    },

    // Section: custom styles + numbering
    {
      children: [
        {
          paragraph: {
            style: "CustomHeading",
            children: ["Custom Heading Style Applied"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "custom-list", level: 0 },
            children: ["First numbered item"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "custom-list", level: 0 },
            children: ["Second numbered item"],
          },
        },
      ],
    },
  ];

  // 2. Create document and export
  const doc = new Document({
    footnotes: {
      "1": { children: ["This is a footnote."] },
    },
    comments: {
      children: [
        {
          id: 0,
          author: "Demo Author",
          initials: "DA",
          children: ["This is a comment via JSON API."],
        },
      ],
    },
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22 },
        },
      },
      paragraphStyles: [
        {
          id: "CustomHeading",
          name: "Custom Heading",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { bold: true, size: 32, color: "1F4E79" },
        },
      ],
      characterStyles: [
        {
          id: "CustomHighlight",
          name: "Custom Highlight",
          run: { bold: true, color: "FF0000" },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: "custom-list",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: "start",
              start: 1,
              suffix: "space",
            },
          ],
        },
      ],
    },
    sections,
  });
  const buffer = await Packer.toBuffer(doc);
  console.log(`Generated DOCX: ${buffer.length} bytes`);

  // 3. Parse it back
  const parsed = parseDocument(buffer);
  console.log(`Parsed ${parsed.sections!.length} sections`);

  // 4. Verify
  console.log("\n--- Verification ---");

  assert("section count = 8", parsed.sections!.length === 8);

  // Section: paragraphs + tables + SDT + textbox + headings + TOC
  const s1 = parsed.sections![0].children;
  assert("s1 has content (>= 12)", s1.length >= 12);

  // Comment paragraph: commentRangeStart, commentRangeEnd, commentReference preserved
  {
    const commentPara = s1.find((c): boolean => {
      if (!("paragraph" in c)) return false;
      const p = (c as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      return children?.some(
        (ch) => typeof ch === "object" && "commentRangeStart" in (ch as Record<string, unknown>),
      );
    });
    assert("s1 has comment paragraph with commentRangeStart", !!commentPara);
  }

  // Hyperlink paragraph: verify link target preserved
  {
    const hlPara = s1.find((c): boolean => {
      if (!("paragraph" in c)) return false;
      const p = (c as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      return children?.some(
        (ch) =>
          typeof ch === "object" && ch !== null && "hyperlink" in (ch as Record<string, unknown>),
      );
    });
    assert("s1 has hyperlink paragraph", !!hlPara);
    if (hlPara) {
      const p = (hlPara as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      const hlChild = children.find(
        (ch) =>
          typeof ch === "object" && ch !== null && "hyperlink" in (ch as Record<string, unknown>),
      ) as Record<string, unknown> | undefined;
      if (hlChild) {
        const hl = hlChild.hyperlink as Record<string, unknown>;
        assert("hyperlink target preserved", hl.link === "https://example.com");
        assert("hyperlink tooltip preserved", hl.tooltip === "Example");
      }
    }
  }

  // Bookmark paragraph: verify bookmark range preserved
  {
    const bmPara = s1.find((c): boolean => {
      if (!("paragraph" in c)) return false;
      const p = (c as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      return children?.some(
        (ch) => typeof ch === "object" && "bookmarkStart" in (ch as Record<string, unknown>),
      );
    });
    assert("s1 has bookmark paragraph", !!bmPara);
    if (bmPara) {
      const p = (bmPara as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      const bmStart = children.find(
        (ch) => typeof ch === "object" && "bookmarkStart" in (ch as Record<string, unknown>),
      ) as Record<string, unknown> | undefined;
      if (bmStart) {
        const bm = bmStart.bookmarkStart as Record<string, unknown>;
        assert("bookmark name preserved", bm.name === "my-bookmark");
        assert("bookmark id preserved", bm.id === 42);
      }
    }
  }
  assert("s1[0] is paragraph", "paragraph" in s1[0]);
  assert(
    "s1 has table",
    s1.some((c) => "table" in c),
  );
  assert(
    "s1 has sdt",
    s1.some((c) => "sdt" in c),
  );
  assert(
    "s1 has textbox",
    s1.some((c) => "textbox" in c),
  );
  assert(
    "s1 has toc",
    s1.some((c) => "toc" in c),
  );

  // Section: chart
  const s2 = parsed.sections![1].children;
  assert("s2 has content (>= 2)", s2.length >= 2);
  {
    const chartPara = s2.find((c): boolean => {
      if (!("paragraph" in c)) return false;
      const p = (c as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      return children?.some((ch) => "chart" in ch);
    });
    assert("s2 has chart child", !!chartPara);
    if (chartPara) {
      const p = (chartPara as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      const chartChild = children.find((ch) => "chart" in ch);
      if (chartChild) {
        const chart = chartChild.chart as Record<string, unknown>;
        assert("chart type = column", chart.type === "column");
        assert("chart title preserved", chart.title === "Quarterly Revenue");
        const data = chart.data as Record<string, unknown>;
        const cats = data.categories as string[];
        assert("chart categories preserved", cats.length === 4 && cats[0] === "Q1");
        const series = data.series as Record<string, unknown>[];
        assert("chart series count = 2", series.length === 2);
      }
    }
  }

  // Section: smartArt
  const s3 = parsed.sections![2].children;
  assert("s3 has content (>= 2)", s3.length >= 2);
  {
    const saPara = s3.find((c): boolean => {
      if (!("paragraph" in c)) return false;
      const p = (c as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      return children?.some((ch) => "smartArt" in ch);
    });
    assert("s3 has smartArt child", !!saPara);
    if (saPara) {
      const p = (saPara as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      const saChild = children.find((ch) => "smartArt" in ch);
      if (saChild) {
        const sa = saChild.smartArt as Record<string, unknown>;
        assert("smartArt layout preserved", sa.layout === "process1");
        assert("smartArt style preserved", sa.style === "simple1");
        assert("smartArt color preserved", sa.color === "accent1_2");
        const data = sa.data as Record<string, unknown>;
        const nodes = data.nodes as Record<string, unknown>[];
        assert("smartArt nodes count = 3", nodes.length === 3);
        assert("smartArt node text preserved", nodes[0].text === "Plan");
        const planChildren = nodes[0].children as Record<string, unknown>[];
        assert(
          "smartArt nested children preserved",
          planChildren?.length === 1 && planChildren[0].text === "Define scope",
        );
      }
    }
  }

  // Section: altChunk
  const s4 = parsed.sections![3].children;
  assert("s4 has content (>= 2)", s4.length >= 2);
  assert(
    "s4 has altChunk",
    s4.some((c) => "altChunk" in c),
  );

  // Section: image
  const s5 = parsed.sections![4].children;
  assert("s5 has content (>= 2)", s5.length >= 2);
  {
    const imgPara = s5.find((c): boolean => {
      if (!("paragraph" in c)) return false;
      const p = (c as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      return children?.some(
        (ch) => typeof ch === "object" && ch !== null && "image" in (ch as Record<string, unknown>),
      );
    });
    assert("s5 has image child", !!imgPara);
    if (imgPara) {
      const p = (imgPara as Record<string, unknown>).paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      const imgChild = children.find(
        (ch) => typeof ch === "object" && ch !== null && "image" in (ch as Record<string, unknown>),
      );
      if (imgChild) {
        const img = (imgChild as Record<string, unknown>).image as Record<string, unknown>;
        assert("image type = png", img.type === "png");
        const t = img.transformation as Record<string, unknown>;
        assert("image width preserved", t?.width === 200);
        assert("image height preserved", t?.height === 200);
      }
    }
  }

  // Section: math, symbol, breaks, footnote
  const s6 = parsed.sections![5].children;
  assert("s6 has content (>= 6)", s6.length >= 6);
  {
    // First paragraph: math + symbolRun
    const first = s6[0] as Record<string, unknown>;
    if ("paragraph" in first) {
      const p = first.paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      assert(
        "s6 has math element parsed as JSON",
        children?.some((ch) => {
          if (typeof ch !== "object" || ch === null) return false;
          const obj = ch as Record<string, unknown>;
          if (!("math" in obj)) return false;
          const math = obj.math as Record<string, unknown>;
          const mathChildren = math.children as Record<string, unknown>[];
          return Array.isArray(mathChildren) && mathChildren.length > 0;
        }),
      );
    }
    // Second paragraph: fraction math
    const second = s6[1] as Record<string, unknown>;
    if ("paragraph" in second) {
      const p = second.paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      assert(
        "s6 fraction math parsed",
        children?.some((ch) => {
          if (typeof ch !== "object" || ch === null || !("math" in ch)) return false;
          const math = (ch as Record<string, unknown>).math as Record<string, unknown>;
          const mathChildren = math.children as Record<string, unknown>[];
          return mathChildren?.some(
            (mc) => typeof mc === "object" && mc !== null && "fraction" in mc,
          );
        }),
      );
    }
    // Third paragraph: sum with superScript
    const third = s6[2] as Record<string, unknown>;
    if ("paragraph" in third) {
      const p = third.paragraph as Record<string, unknown>;
      const children = p.children as Record<string, unknown>[];
      assert(
        "s6 sum math parsed",
        children?.some((ch) => {
          if (typeof ch !== "object" || ch === null || !("math" in ch)) return false;
          const math = (ch as Record<string, unknown>).math as Record<string, unknown>;
          const mathChildren = math.children as Record<string, unknown>[];
          return mathChildren?.some((mc) => typeof mc === "object" && mc !== null && "sum" in mc);
        }),
      );
    }
    // Fourth paragraph: pageBreak
    const fourth = s6[3] as Record<string, unknown>;
    if ("paragraph" in fourth) {
      const p = fourth.paragraph as Record<string, unknown>;
      const pChildren = p.children as Record<string, unknown>[];
      assert("s6 pageBreak paragraph has 3 children", pChildren?.length === 3);
    }
    // Fifth paragraph: columnBreak
    const fifth = s6[4] as Record<string, unknown>;
    if ("paragraph" in fifth) {
      const p = fifth.paragraph as Record<string, unknown>;
      const pChildren = p.children as Record<string, unknown>[];
      assert("s6 columnBreak paragraph has 3 children", pChildren?.length === 3);
    }
    // Sixth paragraph: footnoteReference (parsed as styled run)
    const sixth = s6[5] as Record<string, unknown>;
    if ("paragraph" in sixth) {
      const p = sixth.paragraph as Record<string, unknown>;
      const pChildren = p.children as Record<string, unknown>[];
      assert(
        "s6 has footnoteReference",
        pChildren?.some(
          (ch) =>
            typeof ch === "object" &&
            "style" in (ch as Record<string, unknown>) &&
            (ch as Record<string, unknown>).style === "FootnoteReference",
        ),
      );
    }
  }

  // Section: rich formatting for value comparison
  const s7 = parsed.sections![6].children;
  assert("s7 has >= 1 child", s7.length >= 1);
  if ("paragraph" in s7[0] && typeof s7[0].paragraph === "object" && s7[0].paragraph !== null) {
    const opts = s7[0].paragraph as Record<string, unknown>;
    const children = opts.children as Record<string, unknown>[];
    if (children?.[0]) {
      const run = children[0];
      assert("bold preserved", run.bold === true);
      assert("italics preserved", run.italics === true);
      assert("size preserved", run.size === 28);
      assert("color preserved", run.color === "FF0000");
    }
  }

  // Comments content: verify comment body parsed
  {
    const comments = (parsed as unknown as Record<string, unknown>).comments as
      | Record<string, unknown>
      | undefined;
    assert("comments parsed", !!comments);
    if (comments) {
      const commentChildren = comments.children as Record<string, unknown>[];
      assert("comments has children", Array.isArray(commentChildren) && commentChildren.length > 0);
      if (commentChildren?.[0]) {
        const c = commentChildren[0];
        assert("comment author preserved", c.author === "Demo Author");
        assert("comment id preserved", c.id === 0);
      }
    }
  }

  // Footnotes content: verify footnote body parsed
  {
    const footnotes = (parsed as unknown as Record<string, unknown>).footnotes as
      | Record<string, unknown>
      | undefined;
    assert("footnotes parsed", !!footnotes);
    if (footnotes) {
      const fn1 = (footnotes as Record<string, Record<string, unknown>>)["1"];
      assert("footnote 1 parsed", !!fn1);
      if (fn1) {
        assert("footnote 1 has children", Array.isArray(fn1.children) && fn1.children.length > 0);
      }
    }
  }

  // Styles: verify parsed style definitions
  {
    const styles = (parsed as unknown as Record<string, unknown>).styles as
      | Record<string, unknown>
      | undefined;
    assert("styles parsed", !!styles);
    if (styles) {
      const pStyles = styles.paragraphStyles as Record<string, unknown>[] | undefined;
      assert("paragraphStyles parsed", Array.isArray(pStyles) && pStyles.length > 0);
      if (pStyles && pStyles.length > 0) {
        const customHeading = pStyles.find((s) => s.id === "CustomHeading");
        assert("CustomHeading style found", !!customHeading);
        if (customHeading) {
          assert("CustomHeading name", customHeading.name === "Custom Heading");
          assert("CustomHeading quickFormat", customHeading.quickFormat === true);
        }
      }
      const cStyles = styles.characterStyles as Record<string, unknown>[] | undefined;
      assert("characterStyles parsed", Array.isArray(cStyles) && cStyles.length > 0);
    }
  }

  // Numbering: verify parsed numbering definitions
  {
    const numbering = (parsed as unknown as Record<string, unknown>).numbering as
      | Record<string, unknown>
      | undefined;
    assert("numbering parsed", !!numbering);
    if (numbering) {
      const config = numbering.config as Record<string, unknown>[];
      assert("numbering config parsed", Array.isArray(config) && config.length > 0);
    }
  }

  // Section 8: verify style/numbering paragraphs
  const s8 = parsed.sections![7].children;
  assert("s8 has 3 children", s8.length === 3);
  {
    const headingPara = s8[0] as Record<string, unknown>;
    if ("paragraph" in headingPara) {
      const p = headingPara.paragraph as Record<string, unknown>;
      assert("s8 heading has style", p.style === "CustomHeading");
    }
  }

  // 5. Re-export and re-parse to verify stability
  const doc2 = new Document(parsed);
  const buffer2 = await Packer.toBuffer(doc2);
  const parsed2 = parseDocument(buffer2);
  assert("stable re-parse section count", parsed2.sections!.length === parsed.sections!.length);
  console.log(`Re-exported DOCX: ${buffer2.length} bytes`);

  // Save
  fs.writeFileSync("My Document.docx", buffer);
  fs.writeFileSync("My Document (round-trip).docx", buffer2);

  console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
  if (fail > 0) process.exit(1);
  console.log("\nSaved My Document.docx and My Document (round-trip).docx");
}

main().catch(console.error);
