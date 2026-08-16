// Combines every safe authoring domain in one deck — shapes, rich text,
// media, groups, tables, charts, SmartArt, connectors, animations — to
// surface combination defects the single-feature demos cannot hit.
import { mkdirSync, writeFileSync } from "node:fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const png1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const options: PresentationOptions = {
  title: "Maximum coverage deck",
  creator: "office-open",
  show: { type: "present", loop: false, penColor: "FF0000" },
  view: { lastView: "slideSorterView" },
  slides: [
    // ── Slide 1: shapes + rich text + animation + notes + transition ──
    {
      children: [
        {
          shape: {
            id: 2,
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Maximum coverage deck" },
            geometry: "rect",
            fill: "4472C4",
          },
        },
        {
          shape: {
            id: 3,
            x: "2cm",
            y: "3cm",
            width: "8cm",
            height: "2cm",
            textBody: {
              paragraphs: [
                {
                  children: [
                    { text: "Rich ", bold: true },
                    { text: "text", italic: true, fill: "C00000" },
                    { text: " runs" },
                  ],
                  properties: { bullet: { type: "char", char: "•" }, indentLevel: 0 },
                },
              ],
            },
          },
        },
        {
          line: { id: 4, x1: "2cm", y1: "6cm", x2: "12cm", y2: "6cm" },
        },
      ],
      animations: [{ shapeId: 2, options: { type: "fade", duration: 800 } }],
      transition: { type: "fade", speed: "medium" },
      notes: "Speaker notes for the title slide.",
    },
    // ── Slide 2: picture + hyperlink + group ──
    {
      children: [
        {
          picture: {
            x: "1.3cm",
            y: "2cm",
            width: "6cm",
            height: "4cm",
            data: png1x1,
            type: "png",
            description: "Tiny pixel",
          },
        },
        {
          shape: {
            x: "9cm",
            y: "2cm",
            width: "5cm",
            height: "1.5cm",
            textBody: {
              paragraphs: [
                {
                  children: [
                    {
                      text: "Open example.com",
                      hyperlink: { url: "https://example.com", tooltip: "Go" },
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          group: {
            x: "1.3cm",
            y: "8cm",
            width: "7.9cm",
            height: "5.3cm",
            children: [
              {
                shape: {
                  x: "1.3cm",
                  y: "8cm",
                  width: "3.9cm",
                  height: "2.6cm",
                  textBody: { text: "Grouped A" },
                  geometry: "ellipse",
                },
              },
              {
                shape: {
                  x: "5.3cm",
                  y: "8cm",
                  width: "3.9cm",
                  height: "2.6cm",
                  textBody: { text: "Grouped B" },
                  geometry: "triangle",
                },
              },
            ],
          },
        },
      ],
    },
    // ── Slide 3: table + chart ──
    {
      children: [
        {
          table: {
            x: "1.3cm",
            y: "1.5cm",
            width: "15.9cm",
            height: "5cm",
            rows: [
              {
                cells: [{ text: "Name", fill: "4472C4" }, { text: "Value" }],
              },
              {
                cells: [{ text: "Alpha" }, { text: "42" }],
              },
            ],
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "8cm",
            width: "15.9cm",
            height: "7cm",
            type: "column",
            title: "Coverage",
            categories: ["A", "B", "C"],
            series: [{ name: "S1", values: [1, 5, 3] }],
            showLegend: true,
          },
        },
      ],
    },
    // ── Slide 4: SmartArt + connector + comment ──
    {
      children: [
        {
          smartart: {
            x: "1.3cm",
            y: "2cm",
            width: "10cm",
            height: "6cm",
            layout: "simple1",
            nodes: [{ text: "One" }, { text: "Two" }],
          },
        },
        {
          connector: {
            id: 7,
            x1: "12cm",
            y1: "3cm",
            x2: "16cm",
            y2: "5cm",
          },
        },
      ],
      comments: [
        {
          author: "Max",
          initials: "M",
          text: "Comment on slide four.",
          x: "10cm",
          y: "10cm",
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/36-max-coverage.pptx", buffer);
console.log("written 36-max-coverage.pptx");
