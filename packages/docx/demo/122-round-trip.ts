// Round-trip demo — one document with all content types, export → parse → verify → re-export.

import * as fs from "fs";

import { Document, Packer, parseDocument } from "@office-open/docx";

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
  const sections = [
    // Section 1: paragraphs, tables, nested tables, SDT, headings, TOC, headers/footers
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
            alignment: "center" as const,
            children: [{ text: "Centered text", bold: true, size: 32 }],
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

    // Section 2: chart
    {
      children: [
        { paragraph: { children: [{ text: "Chart (JSON API)", bold: true, size: 28 }] } },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "column" as const,
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

    // Section 3: smartArt
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

    // Section 4: altChunk
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

    // Section 5: rich formatting for value comparison
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Styled", bold: true, italics: true, size: 28, color: "FF0000" }],
          },
        },
      ],
    },
  ];

  // 2. Create document and export
  const doc = new Document({ sections: sections as import("@office-open/docx").ISectionOptions[] });
  const buffer = await Packer.toBuffer(doc);
  console.log(`Generated DOCX: ${buffer.length} bytes`);

  // 3. Parse it back
  const parsed = parseDocument(new Uint8Array(buffer));
  console.log(`Parsed ${parsed.length} sections`);
  console.log("Parsed:", JSON.stringify(parsed, null, 2));

  // 4. Verify
  console.log("\n--- Verification ---");

  assert("section count = 5", parsed.length === 5);

  // Section 1: paragraphs + tables + SDT + textbox + headings + TOC
  const s1 = parsed[0].children;
  assert("s1 has content (>= 12)", s1.length >= 12);
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

  // Section 2: chart
  const s2 = parsed[1].children;
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

  // Section 3: smartArt
  const s3 = parsed[2].children;
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

  // Section 4: altChunk
  const s4 = parsed[3].children;
  assert("s4 has content (>= 2)", s4.length >= 2);
  assert(
    "s4 has altChunk",
    s4.some((c) => "altChunk" in c),
  );

  // Section 5 (last section): value comparison
  const s5 = parsed[4].children;
  assert("s5 has 1 child", s5.length === 1);
  if ("paragraph" in s5[0] && typeof s5[0].paragraph === "object" && s5[0].paragraph !== null) {
    const opts = s5[0].paragraph as Record<string, unknown>;
    const children = opts.children as Record<string, unknown>[];
    if (children?.[0]) {
      const run = children[0];
      assert("bold preserved", run.bold === true);
      assert("italics preserved", run.italics === true);
      assert("size preserved", run.size === 28);
      assert("color preserved", run.color === "FF0000");
    }
  }

  // 5. Re-export and re-parse to verify stability
  const doc2 = new Document({ sections: parsed });
  const buffer2 = await Packer.toBuffer(doc2);
  const parsed2 = parseDocument(new Uint8Array(buffer2));
  assert("stable re-parse section count", parsed2.length === parsed.length);
  console.log(`Re-exported DOCX: ${buffer2.length} bytes`);

  // Save
  fs.writeFileSync("My Document.docx", buffer);
  fs.writeFileSync("My Document (round-trip).docx", buffer2);

  console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
  if (fail > 0) process.exit(1);
  console.log("\nSaved My Document.docx and My Document (round-trip).docx");
}

main().catch(console.error);
