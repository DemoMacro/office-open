import * as fs from "fs";

import {
  generatePresentation,
  parsePresentation,
  parsePptx,
  parseArchive,
} from "@office-open/pptx";
import type { SlideOptions } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";
import { findChild, attrNum, xml2js, js2xml } from "@office-open/xml";
import { strFromU8 } from "fflate";

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

function normalizeXml(raw: Uint8Array): string {
  const parsed = xml2js(strFromU8(raw), {
    nativeTypeAttributes: true,
    captureSpacesBetweenElements: true,
  });
  let xml = js2xml(parsed);
  // Normalize non-deterministic values: cNvPr id, GraphicFrame name/id, diagram UUIDs
  xml = xml.replace(/ id="\d+"/g, ' id="N"');
  xml = xml.replace(/ name="(Table|Chart|Diagram|Group) \d+"/g, ' name="$1 N"');
  xml = xml.replace(/\{[0-9A-F-]{36}\}/g, "{UUID}");
  return xml;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

interface Diff {
  path: string;
  kind: "missing-left" | "missing-right" | "content";
}

function compareZips(buf1: Uint8Array, buf2: Uint8Array, ignorePaths?: Set<string>): Diff[] {
  const zip1 = parseArchive(buf1);
  const zip2 = parseArchive(buf2);
  const diffs: Diff[] = [];

  const allPaths = new Set([...zip1.keys(), ...zip2.keys()]);

  for (const path of allPaths) {
    if (ignorePaths?.has(path)) continue;

    const raw1 = zip1.getRaw(path);
    const raw2 = zip2.getRaw(path);

    if (!raw1) {
      diffs.push({ path, kind: "missing-left" });
      continue;
    }
    if (!raw2) {
      diffs.push({ path, kind: "missing-right" });
      continue;
    }

    if (path.endsWith(".xml") || path.endsWith(".rels")) {
      const s1 = normalizeXml(raw1);
      const s2 = normalizeXml(raw2);
      if (s1 !== s2) {
        diffs.push({ path, kind: "content" });
      }
    } else {
      if (!bytesEqual(raw1, raw2)) {
        diffs.push({ path, kind: "content" });
      }
    }
  }

  return diffs;
}

function printDiffs(diffs: Diff[]): void {
  if (diffs.length === 0) {
    console.log("  No differences found!");
    return;
  }
  for (const d of diffs) {
    const label =
      d.kind === "missing-left"
        ? "only in buffer2"
        : d.kind === "missing-right"
          ? "only in buffer1"
          : "content differs";
    console.log(`  DIFF: ${d.path} (${label})`);
  }
}

// ── media assets ───────────────────────────────────────────────────────────────

const imageData = fs.readFileSync("./demo/assets/test-poster.png");
const videoData = fs.readFileSync("./demo/assets/test-video.mp4");
const posterData = fs.readFileSync("./demo/assets/test-poster.png");

// Generate a minimal valid WAV file (1 second of silence, mono, 8000 Hz)
const wavSize = 8000; // 1 second × 8000 samples
const wavBuf = Buffer.alloc(44 + wavSize);
wavBuf.write("RIFF", 0);
wavBuf.writeUInt32LE(36 + wavSize, 4);
wavBuf.write("WAVE", 8);
wavBuf.write("fmt ", 12);
wavBuf.writeUInt32LE(16, 16); // chunk size
wavBuf.writeUInt16LE(1, 20); // PCM
wavBuf.writeUInt16LE(1, 22); // mono
wavBuf.writeUInt32LE(8000, 24); // sample rate
wavBuf.writeUInt32LE(8000, 28); // byte rate
wavBuf.writeUInt16LE(1, 32); // block align
wavBuf.writeUInt16LE(8, 34); // bits per sample
wavBuf.write("data", 36);
wavBuf.writeUInt32LE(wavSize, 40);
// data is all zeros (silence)

// ══════════════════════════════════════════════════════════════════════════════
// 1. Generate a presentation using the JSON-friendly API
// ══════════════════════════════════════════════════════════════════════════════

const slides: SlideOptions[] = [
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
          textBody: {
            children: [
              {
                properties: { alignment: "center", bullet: { type: "none" } },
                children: [
                  { text: "Round-trip Feature Showcase", size: 32, bold: true, fill: "FFFFFF" },
                ],
              },
            ],
          },
        },
      },
      // Geometry + fill + outline + effects (shadow, glow)
      {
        shape: {
          x: 50,
          y: 110,
          width: 220,
          height: 130,
          textBody: { text: "Shadow + Glow" },
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
          textBody: { text: "Gradient" },
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
          textBody: {
            children: [
              {
                properties: { bullet: { type: "none" } },
                children: [
                  { text: "Bold", bold: true, size: 14 },
                  { text: " " },
                  { text: "Italic", italic: true, size: 14 },
                  { text: "\n" },
                  { text: "Underline", underline: "single", size: 14 },
                  { text: " " },
                  { text: "Strike", strike: "sngStrike", size: 14, fill: "FF0000" },
                ],
              },
            ],
          },
        },
      },
      // Bullet list (char + autoNum)
      {
        shape: {
          x: 50,
          y: 270,
          width: 340,
          height: 160,
          outline: { color: "4472C4", width: 12700, dashStyle: "dash" },
          textBody: {
            children: [
              {
                properties: { bullet: { type: "char", char: "\u25CF", color: "4472C4" } },
                children: [{ text: "Bullet point one" }],
              },
              {
                properties: { bullet: { type: "char", char: "\u25CF", color: "4472C4" } },
                children: [{ text: "Bullet point two" }],
              },
              {
                properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                children: [{ text: "Numbered item" }],
              },
            ],
          },
        },
      },
      // 3D effects (rotation, bevel, extrusion, material)
      {
        shape: {
          x: 550,
          y: 270,
          width: 250,
          height: 100,
          textBody: { text: "3D" },
          geometry: "roundRect",
          fill: "70AD47",
          effects: {
            rotation3D: { x: 20, y: 10, z: 5 },
            bevelTop: { width: 8, height: 8 },
            extrusionH: 30000,
            material: "plastic",
          },
        },
      },
      // Alignment + line spacing + super/subscript
      {
        shape: {
          x: 420,
          y: 390,
          width: 380,
          height: 160,
          textBody: {
            children: [
              {
                properties: { alignment: "center", bullet: { type: "none" }, lineSpacing: 1.5 },
                children: [
                  { text: "H", size: 20 },
                  { text: "2", size: 12, baseline: -25000 },
                  { text: "O  \u2014  E=mc", size: 20 },
                  { text: "2", size: 12, baseline: 30000 },
                ],
              },
              {
                properties: { alignment: "right", bullet: { type: "none" } },
                children: [{ text: "Right-aligned text", fill: "70AD47", font: "Consolas" }],
              },
            ],
          },
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
          textBody: { text: "Table" },
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
            {
              cells: [
                { text: "Shapes" },
                { text: "OK" },
                { text: "rect, roundRect, chevron\u2026" },
              ],
            },
            { cells: [{ text: "Tables" }, { text: "OK" }, { text: "rows, cells, borders" }] },
            { cells: [{ text: "Merge", columnSpan: 2 }, { text: "Spanned cell" }] },
          ],
          columnWidths: [2500000, 2000000, 3000000],
          firstRow: true,
          bandRow: true,
          borders: {
            top: { color: "4472C4", width: 12700 },
            bottom: { color: "4472C4", width: 12700 },
            left: { color: "4472C4", width: 12700 },
            right: { color: "4472C4", width: 12700 },
          },
        },
      },
    ],
    transition: { type: "push", speed: "med", direction: "left" },
  },

  // ── Slide 3: Chart (column) ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "Chart" },
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
          textBody: { text: "SmartArt" },
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
          textBody: { text: "Lines & Connectors" },
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
          textBody: { text: "Group" },
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
            {
              shape: {
                x: 0,
                y: 0,
                width: 290,
                height: 180,
                textBody: { text: "A" },
                fill: "4472C4",
              },
            },
            {
              shape: {
                x: 310,
                y: 0,
                width: 290,
                height: 180,
                textBody: { text: "B" },
                fill: "ED7D31",
              },
            },
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
          textBody: {
            children: [
              {
                properties: { alignment: "center", bullet: { type: "none" } },
                children: [
                  {
                    text: "Gradient Background + Fade Transition",
                    size: 24,
                    bold: true,
                    fill: "FFFFFF",
                  },
                ],
              },
            ],
          },
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
          textBody: { text: "Animation" },
          fill: "C00000",
        },
      },
      {
        shape: {
          id: 2,
          x: 100,
          y: 120,
          width: 250,
          height: 100,
          textBody: { text: "Fly In" },
          fill: "4472C4",
        },
      },
      {
        shape: {
          id: 3,
          x: 400,
          y: 120,
          width: 250,
          height: 100,
          textBody: { text: "Appear" },
          fill: "70AD47",
        },
      },
      {
        shape: {
          id: 4,
          x: 250,
          y: 280,
          width: 250,
          height: 100,
          textBody: { text: "Fade Exit" },
          fill: "ED7D31",
        },
      },
    ],
    animations: [
      { shapeId: 2, options: { type: "fly", class: "entr", direction: "left", duration: 500 } },
      { shapeId: 3, options: { type: "appear", class: "entr", trigger: "afterPrevious" } },
      { shapeId: 4, options: { type: "fade", class: "exit", duration: 750 } },
    ],
  },

  // ── Slide 9: Notes + Header/Footer ──
  {
    notes: "These are speaker notes for the round-trip test.",
    headerFooter: { slideNumber: true, footer: "Round-trip Demo", dateTime: true },
    children: [
      {
        shape: {
          x: 200,
          y: 200,
          width: 400,
          height: 60,
          textBody: { text: "Notes + Header/Footer slide" },
          fill: "F2F2F2",
        },
      },
    ],
  },

  // ── Slide 10: Multi-master (light theme) ──
  {
    master: "light",
    layout: "blank",
    children: [
      {
        shape: {
          x: 100,
          y: 100,
          width: 760,
          height: 400,
          fill: "F2F2F2",
          textBody: { text: "Light theme slide" },
        },
      },
    ],
  },

  // ── Slide 11: Multi-master (dark theme) ──
  {
    master: "dark",
    layout: "blank",
    children: [
      {
        shape: {
          x: 100,
          y: 100,
          width: 760,
          height: 400,
          fill: "2D4A7A",
          textBody: {
            children: [
              {
                properties: { bullet: { type: "none" } },
                children: [{ text: "Dark theme slide", fill: "FFFFFF", size: 24 }],
              },
            ],
          },
        },
      },
    ],
  },

  // ── Slide 12: Picture + Fill Types ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "Picture + Fill Types" },
          fill: "7030A0",
        },
      },
      // Picture element
      {
        picture: {
          x: 50,
          y: 80,
          width: 200,
          height: 200,
          data: imageData,
          type: "png",
          name: "Test Image",
        },
      },
      // Pattern fill
      {
        shape: {
          x: 300,
          y: 80,
          width: 200,
          height: 150,
          fill: {
            type: "pattern",
            pattern: "diagCross",
            foregroundColor: "4472C4",
            backgroundColor: "FFFFFF",
          },
          textBody: { text: "Pattern" },
        },
      },
      // No fill (outline only)
      {
        shape: {
          x: 550,
          y: 80,
          width: 200,
          height: 150,
          fill: { type: "none" },
          outline: { color: "ED7D31", width: 25400 },
          textBody: { text: "No Fill" },
        },
      },
      // Radial gradient
      {
        shape: {
          x: 300,
          y: 280,
          width: 200,
          height: 150,
          fill: {
            type: "gradient",
            path: "circle",
            stops: [
              { position: 0, color: "FFFFFF" },
              { position: 100, color: "4472C4" },
            ],
          },
          textBody: { text: "Radial" },
        },
      },
      // Solid fill with transparency
      {
        shape: {
          x: 550,
          y: 280,
          width: 200,
          height: 150,
          fill: { type: "solid", color: "FF0000" },
          textBody: { text: "Solid" },
        },
      },
    ],
  },

  // ── Slide 13: More Effects (innerShadow, reflection, softEdge) ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "More Effects" },
          fill: "44546A",
        },
      },
      // Inner shadow
      {
        shape: {
          x: 50,
          y: 80,
          width: 250,
          height: 150,
          fill: "FFC000",
          textBody: { text: "Inner Shadow" },
          effects: {
            innerShadow: {
              blur: 40000,
              distance: 30000,
              direction: 5400000,
              color: "000000",
              alpha: 40,
            },
          },
        },
      },
      // Reflection
      {
        shape: {
          x: 350,
          y: 80,
          width: 250,
          height: 150,
          fill: "4472C4",
          textBody: { text: "Reflection" },
          effects: {
            reflection: {
              blurRadius: 6350,
              distance: 38100,
              direction: 5400000,
              startAlpha: 90,
              endAlpha: 0,
            },
          },
        },
      },
      // Soft edge
      {
        shape: {
          x: 50,
          y: 280,
          width: 250,
          height: 150,
          fill: "70AD47",
          textBody: { text: "Soft Edge" },
          effects: { softEdge: { radius: 50800 } },
        },
      },
      // Multiple effects combined
      {
        shape: {
          x: 350,
          y: 280,
          width: 250,
          height: 150,
          fill: "ED7D31",
          geometry: "ellipse",
          textBody: { text: "Combined" },
          effects: {
            outerShadow: {
              blur: 38100,
              distance: 25400,
              direction: 5400000,
              color: "000000",
              alpha: 50,
            },
            reflection: { startAlpha: 80, endAlpha: 0 },
            glow: { radius: 50000, color: "ED7D31", alpha: 40 },
          },
        },
      },
    ],
  },

  // ── Slide 14: More Charts (bar, line, pie) ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 10,
          width: 300,
          height: 30,
          textBody: { text: "Chart Types" },
          fill: "ED7D31",
        },
      },
      {
        chart: {
          x: 20,
          y: 50,
          width: 300,
          height: 180,
          type: "bar",
          title: "Bar Chart",
          categories: ["A", "B", "C"],
          series: [{ name: "Values", values: [10, 20, 30] }],
          style: 2,
        },
      },
      {
        chart: {
          x: 340,
          y: 50,
          width: 300,
          height: 180,
          type: "line",
          title: "Line Chart",
          categories: ["Jan", "Feb", "Mar"],
          series: [{ name: "Sales", values: [5, 15, 10] }],
          showLegend: true,
          style: 2,
        },
      },
      {
        chart: {
          x: 180,
          y: 250,
          width: 300,
          height: 200,
          type: "pie",
          title: "Pie Chart",
          categories: ["Red", "Green", "Blue"],
          series: [{ name: "Share", values: [40, 35, 25] }],
          showLegend: true,
          style: 2,
        },
      },
    ],
  },

  // ── Slide 15: Enhanced Table (rowSpan, verticalAlign, cell margins) ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 400,
          height: 40,
          textBody: { text: "Enhanced Table" },
          fill: "44546A",
        },
      },
      {
        table: {
          x: 50,
          y: 80,
          width: 700,
          height: 300,
          rows: [
            {
              cells: [
                { text: "Header 1", fill: "4472C4", verticalAlign: "center" },
                { text: "Header 2", fill: "4472C4", verticalAlign: "center" },
                { text: "Header 3", fill: "4472C4", verticalAlign: "center" },
              ],
            },
            {
              cells: [
                {
                  text: "RowSpan 2",
                  rowSpan: 2,
                  fill: "E8F0FE",
                  verticalAlign: "center",
                  margins: { left: 100000, right: 100000, top: 50000, bottom: 50000 },
                },
                { text: "Cell B1", fill: "F2F2F2" },
                { text: "Cell C1" },
              ],
            },
            {
              cells: [
                // A2 is merged with A1 (rowSpan)
                { text: "Cell B2", fill: "F2F2F2" },
                { text: "Cell C2" },
              ],
            },
            {
              cells: [{ text: "ColSpan 2", columnSpan: 2, fill: "FFF2CC" }, { text: "Last" }],
            },
          ],
          columnWidths: [2500000, 2500000, 2000000],
          firstRow: true,
          lastRow: true,
          firstCol: true,
          bandRow: true,
          bandCol: true,
          borders: {
            top: { color: "333333", width: 12700 },
            bottom: { color: "333333", width: 12700 },
            left: { color: "333333", width: 12700 },
            right: { color: "333333", width: 12700 },
          },
        },
      },
    ],
  },

  // ── Slide 16: Comments + Hyperlink ──
  {
    comments: [
      { author: "Alice Wang", text: "Great slide!", x: 200, y: 50, date: "2026-05-15T10:00:00Z" },
      { author: "Bob Li", text: "Consider a different color.", x: 400, y: 200, initials: "BL" },
    ],
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "Comments + Hyperlink" },
          fill: "7030A0",
        },
      },
      {
        shape: {
          x: 100,
          y: 100,
          width: 600,
          height: 60,
          textBody: {
            children: [
              {
                properties: { bullet: { type: "none" } },
                children: [
                  { text: "Visit " },
                  {
                    text: "our website",
                    hyperlink: { url: "https://example.com", tooltip: "Example" },
                  },
                  { text: " for more." },
                ],
              },
            ],
          },
        },
      },
    ],
  },

  // ── Slide 17: More Transitions + Emphasis Animation ──
  {
    transition: { type: "cover", direction: "right", speed: "med" },
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 400,
          height: 40,
          textBody: { text: "Transitions + Emphasis" },
          fill: "C00000",
        },
      },
      {
        shape: {
          id: 2,
          x: 50,
          y: 100,
          width: 200,
          height: 100,
          textBody: { text: "Grow" },
          fill: "4472C4",
        },
      },
      {
        shape: {
          id: 3,
          x: 300,
          y: 100,
          width: 200,
          height: 100,
          textBody: { text: "Spin" },
          fill: "70AD47",
        },
      },
      {
        shape: {
          id: 4,
          x: 550,
          y: 100,
          width: 200,
          height: 100,
          textBody: { text: "Color" },
          fill: "ED7D31",
        },
      },
      {
        shape: {
          id: 5,
          x: 300,
          y: 260,
          width: 200,
          height: 100,
          textBody: { text: "Pulse" },
          fill: "7030A0",
        },
      },
    ],
    animations: [
      { shapeId: 2, options: { class: "emph", emphasisType: "growShrink", duration: 800 } },
      { shapeId: 3, options: { class: "emph", emphasisType: "spin", duration: 1000 } },
      {
        shapeId: 4,
        options: { class: "emph", emphasisType: "colorChange", color: "FF0000", duration: 800 },
      },
      { shapeId: 5, options: { class: "emph", emphasisType: "pulse", duration: 500 } },
    ],
  },

  // ── Slide 18: Video ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "Video" },
          fill: "2D4A7A",
        },
      },
      {
        video: {
          x: 100,
          y: 80,
          width: 480,
          height: 270,
          data: videoData,
          type: "mp4",
          name: "Test Video",
          poster: posterData,
          posterType: "png",
        },
      },
    ],
  },

  // ── Slide 19: Audio ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "Audio" },
          fill: "44546A",
        },
      },
      {
        audio: {
          x: 300,
          y: 200,
          width: 120,
          height: 120,
          data: wavBuf,
          type: "wav",
          name: "Test Audio",
        },
      },
    ],
  },

  // ── Slide 20: Geometry showcase ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 300,
          height: 40,
          textBody: { text: "Geometry" },
          fill: "4472C4",
        },
      },
      {
        shape: {
          x: 50,
          y: 80,
          width: 150,
          height: 150,
          textBody: { text: "rect" },
          fill: "4472C4",
        },
      },
      {
        shape: {
          x: 220,
          y: 80,
          width: 150,
          height: 150,
          textBody: { text: "ellipse" },
          fill: "ED7D31",
          geometry: "ellipse",
        },
      },
      {
        shape: {
          x: 390,
          y: 80,
          width: 150,
          height: 150,
          textBody: { text: "diamond" },
          fill: "70AD47",
          geometry: "diamond",
        },
      },
      {
        shape: {
          x: 560,
          y: 80,
          width: 150,
          height: 150,
          textBody: { text: "triangle" },
          fill: "FFC000",
          geometry: "triangle",
        },
      },
      {
        shape: {
          x: 50,
          y: 260,
          width: 150,
          height: 150,
          textBody: { text: "pentagon" },
          fill: "7030A0",
          geometry: "pentagon",
        },
      },
      {
        shape: {
          x: 220,
          y: 260,
          width: 150,
          height: 150,
          textBody: { text: "hexagon" },
          fill: "C00000",
          geometry: "hexagon",
        },
      },
      {
        shape: {
          x: 390,
          y: 260,
          width: 150,
          height: 150,
          textBody: { text: "star" },
          fill: "44546A",
          geometry: "star5",
        },
      },
      {
        shape: {
          x: 560,
          y: 260,
          width: 150,
          height: 150,
          textBody: { text: "plus" },
          fill: "2D4A7A",
          geometry: "plus",
        },
      },
    ],
  },

  // ── Slide 21: Text body options (vertical, anchor, autoFit, wrap, margins) ──
  {
    children: [
      {
        shape: {
          x: 50,
          y: 20,
          width: 400,
          height: 40,
          textBody: { text: "Text Body Options" },
          fill: "44546A",
        },
      },
      // Vertical text
      {
        shape: {
          x: 50,
          y: 80,
          width: 100,
          height: 300,
          fill: "4472C4",
          textBody: { text: "Vertical", vertical: "vert" },
        },
      },
      // Anchor bottom
      {
        shape: {
          x: 180,
          y: 80,
          width: 200,
          height: 300,
          fill: "F2F2F2",
          outline: { color: "4472C4", width: 12700 },
          textBody: { text: "Anchored Bottom", anchor: "bottom" },
        },
      },
      // AutoFit shrink
      {
        shape: {
          x: 410,
          y: 80,
          width: 200,
          height: 100,
          fill: "E8F0FE",
          textBody: { text: "AutoFit Shrink Text To Fit Shape", autoFit: "normal" },
        },
      },
      // Flip horizontal
      {
        shape: {
          x: 650,
          y: 80,
          width: 150,
          height: 100,
          fill: "ED7D31",
          textBody: { text: "Flipped" },
          flipHorizontal: true,
        },
      },
    ],
  },
];

const options: PresentationOptions = {
  title: "Round-trip Feature Showcase",
  subject: "Comprehensive PPTX round-trip test",
  creator: "Parser Demo",
  keywords: "test, round-trip, pptx",
  description: "A presentation covering all supported element types",
  lastModifiedBy: "Test Suite",
  revision: 1,
  show: { loop: true, useTimings: true },
  masters: [
    {
      name: "light",
      background: { fill: "FFFFFF" },
      theme: {
        name: "Light Theme",
        colors: {
          dark1: "333333",
          light1: "FFFFFF",
          dark2: "44546A",
          light2: "E7E6E6",
          accent1: "4472C4",
          accent2: "ED7D31",
          accent3: "70AD47",
          accent4: "FFC000",
          accent5: "7030A0",
          accent6: "C00000",
        },
        fonts: { majorFont: "Segoe UI", minorFont: "Calibri" },
      },
      children: [{ shape: { x: 0, y: 640, width: 960, height: 40, fill: "4472C4" } }],
    },
    {
      name: "dark",
      background: { fill: "1B2A4A" },
      theme: {
        name: "Dark Theme",
        colors: {
          dark1: "FFFFFF",
          light1: "1B2A4A",
          dark2: "E7E6E6",
          light2: "333333",
          accent1: "5B9BD5",
          accent2: "F4B183",
          accent3: "A9D18E",
          accent4: "FFD966",
          accent5: "9B7BB6",
          accent6: "FF6B6B",
        },
        fonts: { majorFont: "Segoe UI Light", minorFont: "Calibri" },
      },
      children: [{ shape: { x: 0, y: 640, width: 960, height: 40, fill: "ED7D31" } }],
    },
  ],
  slides,
};

const buffer = await generatePresentation(options);
console.log(`Generated PPTX: ${buffer.length} bytes`);

// ══════════════════════════════════════════════════════════════════════════════
// 2. Low-level parsePptx verification (ParsedArchive API)
// ══════════════════════════════════════════════════════════════════════════════

console.log("\n--- parsePptx (low-level) ---");
const {
  doc: pptxDoc,
  presentation,
  slides: slidePaths,
  slideMasters,
  slideLayouts,
} = parsePptx(buffer);

const slideCount = 21;
assert(`${slideCount} slide paths`, slidePaths.length === slideCount);
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
assert("has ppt/presentation.xml", pptxDoc.has("ppt/presentation.xml"));
assert("has ppt/theme/theme1.xml", pptxDoc.has("ppt/theme/theme1.xml"));

// ══════════════════════════════════════════════════════════════════════════════
// 3. High-level parsePresentation verification (PresentationOptions API)
// ══════════════════════════════════════════════════════════════════════════════

console.log("\n--- parsePresentation (high-level) ---");
const parsed = parsePresentation(buffer);

assert(`${slideCount} slides parsed`, parsed.slides!.length === slideCount);
assert("title preserved", parsed.title === "Round-trip Feature Showcase");
assert("creator preserved", parsed.creator === "Parser Demo");
assert("size is 16:9", parsed.size === "16:9");

// ══════════════════════════════════════════════════════════════════════════════
// 4. Round-trip: re-generate from parsed data → compare ZIPs
// ══════════════════════════════════════════════════════════════════════════════

console.log("\n--- Round-trip ZIP comparison ---");

const buffer2 = await generatePresentation(parsed);
console.log(`Re-generated PPTX: ${buffer2.length} bytes`);

const ignorePaths = new Set(["docProps/core.xml"]);
const diffs = compareZips(buffer, buffer2, ignorePaths);
printDiffs(diffs);
assert("round-trip ZIPs match", diffs.length === 0);

// ══════════════════════════════════════════════════════════════════════════════
// Summary
// ══════════════════════════════════════════════════════════════════════════════

console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);

fs.writeFileSync("My Presentation.pptx", buffer);
fs.writeFileSync("My Presentation (round-trip).pptx", buffer2);
console.log("\nSaved original and round-trip PPTX files");

if (fail > 0) process.exit(1);
