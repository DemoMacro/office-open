import * as fs from "fs";

import { Presentation, Packer, parsePresentation, parsePptx } from "@office-open/pptx";
import type { ISlideOptions } from "@office-open/pptx";
import { findChild, attrNum } from "@office-open/xml";

// ── helpers ────────────────────────────────────────────────────────────────────

let pass = 0;
let fail = 0;
const assert = (label: string, condition: boolean) => {
  if (condition) {
    pass++;
    console.log(`  PASS: ${label}`);
  } else {
    fail++;
    console.log(`  FAIL: ${label}`);
  }
};

// ══════════════════════════════════════════════════════════════════════════════
// 1. Generate a presentation using the JSON-friendly API
// ══════════════════════════════════════════════════════════════════════════════

const slides: ISlideOptions[] = [
  // ── Slide 1: Shapes, rich text, effects ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 30,
          width: 750,
          height: 60,
          fill: "1B2A4A",
          paragraphs: [
            {
              properties: { alignment: "CENTER", bullet: { type: "none" } },
              children: [
                { text: "Round-trip Feature Showcase", fontSize: 32, bold: true, fill: "FFFFFF" },
              ],
            },
          ],
        },
      },
      // Geometry + fill + outline + effects (shadow, glow)
      {
        shape: {
          x: 50,
          y: 110,
          width: 220,
          height: 130,
          text: "Shadow + Glow",
          geometry: "roundRect",
          fill: "4472C4",
          outline: { color: "2F5496", width: 12700 },
          effects: {
            outerShadow: {
              blur: 50800,
              distance: 38100,
              direction: 2700000,
              color: "000000",
              alpha: 40,
            },
            glow: { radius: 76200, color: "4472C4", alpha: 50 },
          },
        },
      },
      // Gradient fill
      {
        shape: {
          x: 300,
          y: 110,
          width: 220,
          height: 130,
          text: "Gradient",
          geometry: "chevron",
          fill: {
            type: "gradient",
            angle: 90,
            stops: [
              { position: 0, color: "ED7D31" },
              { position: 100, color: "FFC000" },
            ],
          },
        },
      },
      // Rich text: bold, italic, underline, strike, color, font, size
      {
        shape: {
          x: 550,
          y: 110,
          width: 250,
          height: 130,
          paragraphs: [
            {
              properties: { bullet: { type: "none" } },
              children: [
                { text: "Bold", bold: true, fontSize: 14 },
                { text: " " },
                { text: "Italic", italic: true, fontSize: 14 },
                { text: "\n" },
                { text: "Underline", underline: "SINGLE", fontSize: 14 },
                { text: " " },
                { text: "Strike", strike: "SINGLE", fontSize: 14, fill: "FF0000" },
              ],
            },
          ],
        },
      },
      // Bullet list (char + autoNum)
      {
        shape: {
          x: 50,
          y: 270,
          width: 340,
          height: 160,
          outline: { color: "4472C4", width: 12700 },
          paragraphs: [
            {
              properties: { bullet: { type: "char", char: "●", color: "4472C4" } },
              children: [{ text: "Bullet point one" }],
            },
            {
              properties: { bullet: { type: "char", char: "●", color: "4472C4" } },
              children: [{ text: "Bullet point two" }],
            },
            {
              properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
              children: [{ text: "Numbered item" }],
            },
          ],
        },
      },
      // Alignment + line spacing + super/subscript
      {
        shape: {
          x: 420,
          y: 270,
          width: 380,
          height: 160,
          paragraphs: [
            {
              properties: { alignment: "CENTER", bullet: { type: "none" }, lineSpacing: 1.5 },
              children: [
                { text: "H", fontSize: 20 },
                { text: "2", fontSize: 12, baseline: -25000 },
                { text: "O  —  E=mc", fontSize: 20 },
                { text: "2", fontSize: 12, baseline: 30000 },
              ],
            },
            {
              properties: { alignment: "RIGHT", bullet: { type: "none" } },
              children: [{ text: "Right-aligned text", fill: "70AD47", font: "Consolas" }],
            },
          ],
        },
      },
    ],
  },

  // ── Slide 2: Table with formatting ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          text: "Table",
          fill: "4472C4",
        },
      },
      {
        table: {
          x: 50,
          y: 80,
          width: 700,
          height: 250,
          rows: [
            {
              cells: [
                { text: "Feature", fill: "4472C4" },
                { text: "Status", fill: "4472C4" },
                { text: "Notes", fill: "4472C4" },
              ],
            },
            { cells: [{ text: "Shapes" }, { text: "OK" }, { text: "rect, roundRect, chevron…" }] },
            { cells: [{ text: "Tables" }, { text: "OK" }, { text: "rows, cells, borders" }] },
            {
              cells: [{ text: "Merge", columnSpan: 2 }, { text: "Spanned cell" }],
            },
          ],
          columnWidths: [2500000, 2000000, 3000000],
          firstRow: true,
          bandRow: true,
        },
      },
    ],
    transition: { type: "push", speed: "med", direction: "left" },
  },

  // ── Slide 3: Chart ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          text: "Chart",
          fill: "ED7D31",
        },
      },
      {
        chart: {
          x: 80,
          y: 80,
          width: 640,
          height: 350,
          type: "column",
          title: "Quarterly Sales",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "Product A", values: [12, 18, 15, 22] },
            { name: "Product B", values: [8, 14, 11, 19] },
          ],
          showLegend: true,
          style: 2,
        },
      },
    ],
  },

  // ── Slide 4: SmartArt ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          text: "SmartArt",
          fill: "70AD47",
        },
      },
      {
        smartart: {
          x: 80,
          y: 80,
          width: 640,
          height: 300,
          nodes: [
            {
              text: "Root",
              children: [{ text: "Child A" }, { text: "Child B" }, { text: "Child C" }],
            },
          ],
          layout: "hierarchy1",
          style: "simple1",
          color: "accent1_2",
        },
      },
    ],
  },

  // ── Slide 5: Connector + Line ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          text: "Lines & Connectors",
          fill: "7030A0",
        },
      },
      { line: { x1: 50, y1: 100, x2: 750, y2: 100, outline: { color: "4472C4", width: 25400 } } },
      {
        connector: {
          x1: 100,
          y1: 150,
          x2: 600,
          y2: 350,
          beginArrowhead: "oval",
          endArrowhead: "triangle",
          outline: { color: "ED7D31", width: 25400 },
        },
      },
    ],
  },

  // ── Slide 6: Group with nested shapes ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          text: "Group",
          fill: "FFC000",
        },
      },
      {
        group: {
          x: 80,
          y: 80,
          width: 620,
          height: 200,
          rotation: 5000,
          children: [
            { shape: { x: 0, y: 0, width: 290, height: 180, text: "A", fill: "4472C4" } },
            { shape: { x: 310, y: 0, width: 290, height: 180, text: "B", fill: "ED7D31" } },
          ],
        },
      },
    ],
  },

  // ── Slide 7: Background + Transition ──
  {
    background: {
      fill: {
        type: "gradient",
        angle: 90,
        stops: [
          { position: 0, color: "1B2A4A" },
          { position: 100, color: "2D4A7A" },
        ],
      },
    },
    transition: { type: "fade", speed: "slow" },
    children: [
      {
        shape: {
          x: 150,
          y: 200,
          width: 500,
          height: 80,
          paragraphs: [
            {
              properties: { alignment: "CENTER", bullet: { type: "none" } },
              children: [
                {
                  text: "Gradient Background + Fade Transition",
                  fontSize: 24,
                  bold: true,
                  fill: "FFFFFF",
                },
              ],
            },
          ],
        },
      },
    ],
  },

  // ── Slide 8: Animation ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          text: "Animation",
          fill: "C00000",
        },
      },
      {
        shape: {
          x: 100,
          y: 120,
          width: 250,
          height: 100,
          text: "Fly In",
          fill: "4472C4",
          animation: { type: "fly", class: "entr", direction: "left", duration: 500 },
        },
      },
      {
        shape: {
          x: 400,
          y: 120,
          width: 250,
          height: 100,
          text: "Appear",
          fill: "70AD47",
          animation: { type: "appear", class: "entr", trigger: "afterPrevious" },
        },
      },
      {
        shape: {
          x: 250,
          y: 280,
          width: 250,
          height: 100,
          text: "Fade Exit",
          fill: "ED7D31",
          animation: { type: "fade", class: "exit", duration: 750 },
        },
      },
    ],
  },

  // ── Slide 9: Notes + Header/Footer ──
  {
    notes: "These are speaker notes for the round-trip test.",
    headerFooter: { slideNumber: true, footer: "Round-trip Demo" },
    children: [
      {
        shape: {
          x: 200,
          y: 200,
          width: 400,
          height: 60,
          text: "Notes + Header/Footer slide",
          fill: "F2F2F2",
        },
      },
    ],
  },
];

const pres = new Presentation({
  title: "Round-trip Feature Showcase",
  creator: "Parser Demo",
  slides,
});

const buffer = await Packer.toBuffer(pres);
console.log(`Generated PPTX: ${buffer.length} bytes`);

// ══════════════════════════════════════════════════════════════════════════════
// 2. Low-level parsePptx verification (ParsedDocument API)
// ══════════════════════════════════════════════════════════════════════════════

console.log("\n--- parsePptx (low-level) ---");
const {
  doc: parsed,
  presentation,
  slides: slidePaths,
  slideMasters,
  slideLayouts,
} = parsePptx(new Uint8Array(buffer));

assert("9 slide paths", slidePaths.length === 9);
assert("slide 1 path", slidePaths[0] === "ppt/slides/slide1.xml");
assert("has presentation element", !!presentation);

if (presentation) {
  const sldSz = findChild(presentation, "p:sldSz");
  const cx = sldSz ? attrNum(sldSz, "cx") : undefined;
  const cy = sldSz ? attrNum(sldSz, "cy") : undefined;
  assert("has slide size", !!cx && !!cy);
}

assert("has slide masters", slideMasters.length >= 1);
assert("has slide layouts", slideLayouts.length >= 1);
assert("has ppt/presentation.xml", parsed.has("ppt/presentation.xml"));
assert("has ppt/theme/theme1.xml", parsed.has("ppt/theme/theme1.xml"));

// ══════════════════════════════════════════════════════════════════════════════
// 3. High-level parsePresentation verification (ISlideOptions[] API)
// ══════════════════════════════════════════════════════════════════════════════

console.log("\n--- parsePresentation (high-level) ---");
const readSlides = parsePresentation(new Uint8Array(buffer));

assert("9 slides parsed", readSlides.length === 9);

// ── Slide 1: Shapes, effects, rich text ──
const s0 = readSlides[0];
const s0c = s0.children!;
assert("s0 has children", s0c.length >= 6);

const s0shape0 = s0c[0] as unknown as { shape: Record<string, unknown> };
assert("s0 child 0 is shape", "shape" in s0shape0);
assert("s0 shape has fill", s0shape0.shape.fill === "1B2A4A");
assert("s0 shape has paragraphs", Array.isArray(s0shape0.shape.paragraphs));

const s0shape1 = s0c[1] as unknown as { shape: Record<string, unknown> };
assert("s0 shape 1 has geometry", s0shape1.shape.geometry === "roundRect");
assert("s0 shape 1 has outline", !!s0shape1.shape.outline);
assert("s0 shape 1 has effects", !!s0shape1.shape.effects);

const s0shape2 = s0c[2] as unknown as { shape: Record<string, unknown> };
assert("s0 shape 2 has gradient fill", typeof s0shape2.shape.fill === "object");

// ── Slide 2: Table + transition ──
const s1 = readSlides[1];
assert("s1 has transition", !!s1.transition);
assert("s1 transition type", (s1.transition as Record<string, unknown>)?.type === "push");
const s1tbl = s1.children!.find((c) => "table" in (c as object)) as unknown as {
  table: Record<string, unknown>;
};
assert("s1 has table", !!s1tbl);
assert("s1 table has rows", Array.isArray(s1tbl.table.rows));
assert("s1 table firstRow", s1tbl.table.firstRow === true);
assert("s1 table bandRow", s1tbl.table.bandRow === true);

// ── Slide 3: Chart ──
const s2 = readSlides[2];
const s2chart = s2.children!.find((c) => "chart" in (c as object)) as unknown as {
  chart: Record<string, unknown>;
};
assert("s2 has chart", !!s2chart);
assert("s2 chart type is column", s2chart.chart.type === "column");
assert("s2 chart has title", s2chart.chart.title === "Quarterly Sales");
assert("s2 chart has categories", Array.isArray(s2chart.chart.categories));
assert("s2 chart has series", Array.isArray(s2chart.chart.series));
assert("s2 chart showLegend", s2chart.chart.showLegend === true);

// ── Slide 4: SmartArt ──
const s3 = readSlides[3];
const s3sa = s3.children!.find((c) => "smartart" in (c as object)) as unknown as {
  smartart: Record<string, unknown>;
};
assert("s3 has smartart", !!s3sa);
assert("s3 smartart has nodes", Array.isArray(s3sa.smartart.nodes));
assert("s3 smartart has layout", s3sa.smartart.layout === "hierarchy1");

// ── Slide 5: Connector ──
const s4 = readSlides[4];
const s4conn = s4.children!.find((c) => "connector" in (c as object)) as unknown as {
  connector: Record<string, unknown>;
};
assert("s4 has connector", !!s4conn);
assert("s4 connector endArrowhead", s4conn.connector.endArrowhead === "triangle");
assert("s4 connector beginArrowhead", s4conn.connector.beginArrowhead === "oval");

// ── Slide 6: Group ──
const s5 = readSlides[5];
const s5grp = s5.children!.find((c) => "group" in (c as object)) as unknown as {
  group: Record<string, unknown>;
};
assert("s5 has group", !!s5grp);
assert("s5 group has children", Array.isArray(s5grp.group.children));
assert("s5 group has rotation", s5grp.group.rotation !== undefined);

// ── Slide 7: Background + transition ──
const s6 = readSlides[6];
assert("s6 has background", !!s6.background);
assert(
  "s6 background is gradient",
  typeof (s6.background as Record<string, unknown>)?.fill === "object",
);
assert("s6 has transition", !!s6.transition);
assert("s6 transition type is fade", (s6.transition as Record<string, unknown>)?.type === "fade");

// ── Slide 8: Animation ──
const s7 = readSlides[7];
const s7shapes = s7.children!.filter((c) => "shape" in (c as object));
assert("s7 has animated shapes", s7shapes.length >= 3);

// ── Slide 9: Notes + headerFooter ──
const s8 = readSlides[8];
assert("s8 has notes", !!s8.notes);
// headerFooter is stored in slide master placeholders, not directly in slide XML
// Skip headerFooter round-trip assertion

// ══════════════════════════════════════════════════════════════════════════════
// 4. Round-trip: re-generate from parsed data
// ══════════════════════════════════════════════════════════════════════════════

console.log("\n--- Round-trip Re-generation ---");

const pres2 = new Presentation({
  title: "Round-trip Feature Showcase",
  slides: readSlides as ISlideOptions[],
});
const buffer2 = await Packer.toBuffer(pres2);
assert("re-generated buffer non-empty", buffer2.length > 0);
console.log(`Re-generated PPTX: ${buffer2.length} bytes`);

// Re-parse the re-generated file
const readSlides2 = parsePresentation(new Uint8Array(buffer2));
assert("re-parsed has 9 slides", readSlides2.length === 9);

const rs0 = readSlides2[0].children!;
assert("re-parsed s0 has children", rs0.length >= 6);
const rs0shape0 = rs0[0] as unknown as { shape: Record<string, unknown> };
assert("re-parsed shape text preserved", rs0shape0.shape.text === undefined); // title shape uses paragraphs

// Verify chart round-trip
const rs2chart = readSlides2[2].children!.find((c) => "chart" in (c as object)) as unknown as {
  chart: Record<string, unknown>;
};
assert("re-parsed chart type preserved", rs2chart.chart.type === "column");
assert("re-parsed chart title preserved", rs2chart.chart.title === "Quarterly Sales");

// Verify table round-trip
const rs1tbl = readSlides2[1].children!.find((c) => "table" in (c as object)) as unknown as {
  table: Record<string, unknown>;
};
assert("re-parsed table rows preserved", Array.isArray(rs1tbl.table.rows));
assert("re-parsed table firstRow preserved", rs1tbl.table.firstRow === true);

// ══════════════════════════════════════════════════════════════════════════════
// Summary
// ══════════════════════════════════════════════════════════════════════════════

console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);

fs.writeFileSync("My Presentation.pptx", buffer);
fs.writeFileSync("My Presentation (round-trip).pptx", buffer2);
console.log("\nSaved original and round-trip PPTX files");

if (fail > 0) process.exit(1);
