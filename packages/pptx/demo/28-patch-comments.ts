import { mkdirSync, writeFileSync } from "node:fs";

import { generatePresentation, patchPresentation } from "@office-open/pptx";

// Patch comments onto existing slides. Authors are merged into
// commentAuthors.xml and per-slide comments into ppt/comments/commentN.xml.
const buffer = await generatePresentation({
  title: "Comment Patch Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2.1cm",
            y: "3.2cm",
            width: "19.1cm",
            height: "2.1cm",
            textBody: { paragraphs: [{ children: [{ text: "Slide one" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "2.1cm",
            y: "3.2cm",
            width: "19.1cm",
            height: "2.1cm",
            textBody: { paragraphs: [{ children: [{ text: "Slide two" }] }] },
          },
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/28-patch-comments.pptx", buffer);

const patched = await patchPresentation({
  outputType: "nodebuffer",
  data: buffer,
  commands: [
    {
      type: "insert",
      at: { field: "slides", route: [{ index: 0 }, { collection: "comments" }] },
      value: [{ author: "Alice", text: "Review the opening", x: "2.6cm", y: "4.0cm" }],
    },
    {
      type: "insert",
      at: { field: "slides", route: [{ index: 1 }, { collection: "comments" }] },
      value: [{ author: "Bob", text: "Add a chart here", x: "2.6cm", y: "4.0cm" }],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/28-patch-comments.pptx", patched);
