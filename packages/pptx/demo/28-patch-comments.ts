import { writeFileSync } from "node:fs";

import { generatePresentation, patchPresentation } from "@office-open/pptx";

// Patch comments onto existing slides. comments keys are 0-based slide indices
// (sldIdLst order). Authors are merged into commentAuthors.xml (deduped by name,
// ids continued); per-slide comments are merged into ppt/comments/commentN.xml.
const buffer = await generatePresentation({
  title: "Comment Patch Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: 80,
            y: 120,
            width: 720,
            height: 80,
            textBody: { children: [{ children: [{ text: "Slide one" }] }] },
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: 80,
            y: 120,
            width: 720,
            height: 80,
            textBody: { children: [{ children: [{ text: "Slide two" }] }] },
          },
        },
      ],
    },
  ],
});

writeFileSync("My Presentation.pptx", buffer);

const patched = await patchPresentation({
  outputType: "nodebuffer",
  data: buffer,
  comments: {
    0: [{ author: "Alice", text: "Review the opening", x: 100, y: 150 }],
    1: [{ author: "Bob", text: "Add a chart here", x: 100, y: 150 }],
  },
});

writeFileSync("My Presentation.pptx", patched);
