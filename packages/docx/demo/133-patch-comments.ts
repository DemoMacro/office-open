import { writeFileSync } from "node:fs";

import { generateDocument, patchDocument } from "@office-open/docx";

// Patch comments onto an existing document. `paragraphs` anchors a comment to
// the Nth body paragraph (0-based, wraps the whole paragraph); `placeholders`
// anchors to the run containing a {{key}} (wrapped before substitution). Comment
// ids continue from any existing word/comments.xml; entries are merged in.
const buffer = await generateDocument({
  title: "Comment Patch Demo",
  sections: [
    {
      children: [
        { paragraph: { children: ["First paragraph of the document."] } },
        { paragraph: { children: ["Second paragraph with a {{name}} placeholder."] } },
      ],
    },
  ],
});

writeFileSync("My Document.docx", buffer);

const patched = await patchDocument({
  outputType: "nodebuffer",
  data: buffer,
  placeholders: { name: { type: "paragraph", children: ["reviewer"] } },
  comments: {
    paragraphs: {
      0: [
        {
          author: "Alice",
          children: [{ children: ["Comment anchored to the first paragraph."] }],
        },
      ],
    },
    placeholders: {
      name: [
        {
          author: "Bob",
          children: [{ children: ["Comment anchored to the placeholder run."] }],
        },
      ],
    },
  },
});

writeFileSync("My Document.docx", patched);
