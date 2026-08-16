// Patch a presentation: remove slides from a template deck.

import { mkdirSync, writeFileSync } from "node:fs";

import { generatePresentation, patchPresentation } from "@office-open/pptx";

// Step 1: Create a four-slide template — a reusable deck with placeholder
//         slides that not every report needs.
const templateBuffer = await generatePresentation({
  title: "Patch Slide Removal Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2cm",
            y: "3cm",
            width: "20cm",
            height: "2cm",
            textBody: { paragraphs: [{ children: [{ text: "Quarterly Report" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "2cm",
            y: "3cm",
            width: "20cm",
            height: "2cm",
            textBody: { paragraphs: [{ children: [{ text: "Appendix A (optional)" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "2cm",
            y: "3cm",
            width: "20cm",
            height: "2cm",
            textBody: { paragraphs: [{ children: [{ text: "Appendix B (optional)" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "2cm",
            y: "3cm",
            width: "20cm",
            height: "2cm",
            textBody: { paragraphs: [{ children: [{ text: "Q&A" }] }] },
          },
        },
      ],
    },
  ],
});

// Step 2: Patch — drop the two appendix slides by their 0-based indices.
//         Removal unwires the whole chain: sldIdLst entries, presentation
//         relationships, slide parts with their rels, and content-type
//         Overrides.
const patched = await patchPresentation({
  outputType: "nodebuffer",
  data: templateBuffer,
  slides: { remove: [1, 2] },
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/37-patch-slide-removal.pptx", patched);
