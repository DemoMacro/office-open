// Add images to header and footer

import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
      footers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  image: {
                    data: readFileSync("./demo/images/pizza.gif"),
                    transformation: {
                      height: 100,
                      width: 100,
                    },
                    type: "gif",
                  },
                },
              ],
            },
          },
        ],
      },
      headers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  image: {
                    data: readFileSync("./demo/images/pizza.gif"),
                    transformation: {
                      height: 100,
                      width: 100,
                    },
                    type: "gif",
                  },
                },
              ],
            },
          },
        ],
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
