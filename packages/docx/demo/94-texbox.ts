// Simple example to add textbox to a document
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          textbox: {
            alignment: "center",
            children: [
              {
                paragraph: {
                  children: ["Hi i'm a textbox!"],
                },
              },
            ],
            style: {
              height: "auto",
              width: "200pt",
            },
          },
        },
        {
          textbox: {
            alignment: "center",
            children: [
              {
                paragraph: {
                  children: ["Hi i'm a textbox with a hidden box!"],
                },
              },
            ],
            style: {
              height: "10.6cm",
              visibility: "hidden",
              width: "300pt",
              zIndex: "auto",
            },
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
