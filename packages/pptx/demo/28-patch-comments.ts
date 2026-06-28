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
            x: "2.1cm",
            y: "3.2cm",
            width: "19.1cm",
            height: "2.1cm",
            textBody: { children: [{ children: [{ text: "Slide one" }] }] },
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
    0: [{ author: "Alice", text: "Review the opening", x: "2.6cm", y: "4.0cm" }],
    1: [{ author: "Bob", text: "Add a chart here", x: "2.6cm", y: "4.0cm" }],
  },
});

writeFileSync("My Presentation.pptx", patched);
