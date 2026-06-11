import { writeFileSync } from "node:fs";
// Example of how to set the document page sizes
// Reference from https://papersizes.io/a/a3

import { PageOrientation, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
      properties: {
        page: {
          size: {
            height: "210mm",
            orientation: PageOrientation.LANDSCAPE,
            width: "148mm",
          },
        },
      },
    },
    {
      children: [{ paragraph: "Hello World" }],
      properties: {
        page: {
          size: {
            height: "420mm",
            orientation: PageOrientation.PORTRAIT,
            width: "297mm",
          },
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
