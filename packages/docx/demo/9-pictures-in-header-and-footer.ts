// Add images to header and footer

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

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
                  picture: {
                    data: readFileSync("./demo/images/pizza.gif"),
                    transformation: {
                      height: "2.6cm",
                      width: "2.6cm",
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
                  picture: {
                    data: readFileSync("./demo/images/pizza.gif"),
                    transformation: {
                      height: "2.6cm",
                      width: "2.6cm",
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
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/9-pictures-in-header-and-footer.docx", buffer);
